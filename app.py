# -*- coding: utf-8 -*-
"""
Excel → Streamlit analyzátor metrik:
- PŘÍJEM (Nákup + Vložit): ST=1, jinak Qty
- EXPEDICE (Vydat + Prodej/prázdné): ST=1, jinak Qty
- TÓNOVÁNÍ (Vložit + Výroba/PO_Pozn): PAL=Qty*24, jinak Qty
- TRANSFERY (Vložit + WorkClass prázdné): počet řádků; NEZAPOČÍTÁVAT do součtů
Výstupy:
  1) Den × Uživatel – metriky + CELKEM (bez Transfery)
  2) Průřez SKU (Den × Uživatel × SKU) pro metriky bez Transfery (+ pivot + Top-N)
  3) Denní součty všichni (bez Transfery)
  4) Transfery zvlášť (Den × LocationBucket)
Grafy: Heatmapa Den×Uživatel, Stacked bar po dnech, Trend, Top-N SKU.
Export: Excel se 4 listy.
"""

import io
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from datetime import date

# -------------------------------
# Nastavení stránky & helpery UI
# -------------------------------
st.set_page_config(page_title="Řádky práce – denní metriky", layout="wide")
st.title("📦 Denní metriky skladových prací (Příjem / Expedice / Tónování / Transfery)")

# -------------------------------
# Helper funkce (normalizace, load)
# -------------------------------

def normalize_str_series(s: pd.Series) -> pd.Series:
    """Odstraní diakritiku, ořízne, upper-case; bezpečné pro NaN."""
    s = s.astype("string").fillna("").str.strip()
    # odstranění diakritiky (NFKD → ascii)
    s = s.str.normalize("NFKD").str.encode("ascii", "ignore").str.decode("ascii")
    return s.str.upper()

def to_excel_date_series(s: pd.Series) -> pd.Series:
    """Excel serial number → pandas datetime.date (ořez na den). Ignoruje NaN."""
    # Excel origin 1899-12-30; pandas zvládne i float serial (část dne)
    dt = pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
    return dt.dt.floor("D").dt.date

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Načte první list nebo CSV a vrátí DataFrame."""
    if filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    elif filename.lower().endswith(".csv"):
        # pokusíme se odhadnout oddělovač, default ;
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=";")
            if df.shape[1] == 1:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=",")
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=",")
    else:
        raise ValueError("Nepodporovaný formát souboru. Použijte XLSX/XLSM/XLS/CSV.")
    return df

def ensure_columns(df: pd.DataFrame, required: list):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Chybí povinné sloupce: {missing}")

def add_computed_columns(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Přidá normalizované sloupce a mezivýpočty metrik."""
    df = df_raw.copy()

    # Povinné sloupce
    required_cols = [
        "Typ práce", "ID pracovní třídy", "Množství práce", "Jednotka",
        "Uzavřená práce", "ID uživatele", "Č. položky", "Místo"
    ]
    ensure_columns(df, required_cols)

    # Normalizace textů pro porovnávání
    df["WorkType_n"]   = normalize_str_series(df["Typ práce"])          # VLOŽIT / VYDAT ...
    df["WorkClass_n"]  = normalize_str_series(df["ID pracovní třídy"])   # NAKUP / PRODEJ / VYROBA / PO_POZN / ''
    df["Unit_n"]       = normalize_str_series(df["Jednotka"])           # ST / PAL / ...
    df["User"]         = df["ID uživatele"].astype("string").fillna("").str.strip()
    df["SKU"]          = df["Č. položky"].astype("string").fillna("").str.strip()
    df["Location"]     = df["Místo"].astype("string").fillna("").str.strip()

    # Den z Excel serial date
    df["Den"] = to_excel_date_series(df["Uzavřená práce"])
    # Množství jako číslo
    df["Qty"] = pd.to_numeric(df["Množství práce"], errors="coerce").fillna(0.0)

    # Metriky
    # Příjem/Expedice: ST -> 1, jinak Qty
    df["EffectiveQty_STis1"] = np.where(df["Unit_n"] == "ST", 1.0, df["Qty"])

    # Tónování: PAL -> Qty * 24, jinak Qty (24 je fixní)
    df["SkpQty"] = np.where(df["Unit_n"] == "PAL", df["Qty"] * 24.0, df["Qty"])

    # Masky metrik (case-insensitive přes normalizaci)
    df["is_prijem"]    = (df["WorkType_n"] == "VLOZIT") & (df["WorkClass_n"] == "NAKUP")
    df["is_expedice"]  = (df["WorkType_n"] == "VYDAT") & ((df["WorkClass_n"] == "PRODEJ") | (df["WorkClass_n"] == ""))
    df["is_tonovani"]  = (df["WorkType_n"] == "VLOZIT") & (df["WorkClass_n"].isin(["VYROBA", "PO_POZN"]))
    df["is_transfer"]  = (df["WorkType_n"] == "VLOZIT") & (df["WorkClass_n"] == "")

    # Bucket pro Transfery: "Adresy (kódy)" vs. textové lokace
    # Vzor adresového kódu: X-1-2-3; podporuje i české znaky v první části (pro jistotu)
    addr_regex = r"^[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ]-\d+-\d+-\d+$"
    df["LocationBucket"] = np.where(
        df["Location"].str.fullmatch(addr_regex, na=False),
        "Adresy (kódy)",
        df["Location"]
    )

    # Ošetření neplatných dnů (NaT → zahodíme)
    df = df[~pd.isna(df["Den"])].copy()

    return df

# -------------------------------
# Načtení dat
# -------------------------------
with st.sidebar:
    st.header("Nahrát data")
    up = st.file_uploader("Excel/CSV s řádky práce", type=["xlsx", "xlsm", "xls", "csv"])
    show_raw_preview = st.checkbox("Zobrazit náhled syrových dat (prvních 200 řádků)", value=False)

if up is None:
    st.info("⬅️ Nahraj prosím soubor v levém panelu. Podporováno: XLSX/XLSM/XLS/CSV.")
    st.stop()

# Load + enrich
try:
    df_src = load_excel(up.read(), up.name)
    df = add_computed_columns(df_src)
except Exception as e:
    st.error(f"Chyba při načítání/zpracování: {e}")
    st.stop()

if show_raw_preview:
    st.subheader("Náhled syrových dat")
    st.dataframe(df_src.head(200), use_container_width=True)

# -------------------------------
# Globální filtry (Den, Uživatel, SKU text)
# -------------------------------
st.sidebar.header("Filtry")
den_min = df["Den"].min()
den_max = df["Den"].max()
default_from = max(den_min, den_max)  # default 1 den = poslední
date_from, date_to = st.sidebar.date_input(
    "Rozsah dní (Uzavřená práce)", value=(default_from, den_max), min_value=den_min, max_value=den_max
) if den_min and den_max else (None, None)

users_all = sorted(df["User"].unique())
sel_users = st.sidebar.multiselect("ID uživatele (login)", options=users_all, default=users_all)

sku_query = st.sidebar.text_input("Filtrovat SKU (obsahuje…)", value="").strip()

# Aplikace filtrů do dat
mask_range = True
if isinstance(date_from, date) and isinstance(date_to, date):
    mask_range = (df["Den"] >= date_from) & (df["Den"] <= date_to)

mask_user = df["User"].isin(sel_users) if sel_users else True
mask_sku  = df["SKU"].str.contains(sku_query, case=False, na=False) if sku_query else True

df_f = df.loc[mask_range & mask_user & mask_sku].copy()

# -------------------------------
# Přehled výpočtů (agregace)
# -------------------------------

def agg_metric(df_in: pd.DataFrame, mask_col: str, value_col: str, by_cols: list) -> pd.DataFrame:
    """Sečte value_col přes masku mask_col dle skupin by_cols."""
    tmp = df_in.loc[df_in[mask_col]].groupby(by_cols, dropna=False)[value_col].sum().reset_index()
    return tmp

# 1) Den × Uživatel – souhrn metrik (bez Transfery)
gcols = ["Den", "User"]
prijem_du   = agg_metric(df_f, "is_prijem",   "EffectiveQty_STis1", gcols).rename(columns={"EffectiveQty_STis1": "Prijem"})
expedice_du = agg_metric(df_f, "is_expedice", "EffectiveQty_STis1", gcols).rename(columns={"EffectiveQty_STis1": "Expedice"})
ton_du      = agg_metric(df_f, "is_tonovani", "SkpQty",             gcols).rename(columns={"SkpQty": "Tonovani"})

den_user = pd.merge(pd.merge(prijem_du, expedice_du, on=gcols, how="outer"),
                    ton_du, on=gcols, how="outer").fillna(0.0)
den_user["Celkem_bez_Transfery"] = den_user[["Prijem", "Expedice", "Tonovani"]].sum(axis=1)

# 2) Průřez SKU – Den × Uživatel × SKU (bez Transfery)
gcols_sku = ["Den", "User", "SKU"]
prijem_sku   = agg_metric(df_f, "is_prijem",   "EffectiveQty_STis1", gcols_sku).rename(columns={"EffectiveQty_STis1": "Prijem"})
expedice_sku = agg_metric(df_f, "is_expedice", "EffectiveQty_STis1", gcols_sku).rename(columns={"EffectiveQty_STis1": "Expedice"})
ton_sku      = agg_metric(df_f, "is_tonovani", "SkpQty",             gcols_sku).rename(columns={"SkpQty": "Tonovani"})

sku_pivot = pd.merge(pd.merge(prijem_sku, expedice_sku, on=gcols_sku, how="outer"),
                     ton_sku, on=gcols_sku, how="outer").fillna(0.0)

# Dlouhý formát na přání (Metrika, Hodnota)
sku_long = sku_pivot.melt(id_vars=gcols_sku, value_vars=["Prijem", "Expedice", "Tonovani"],
                          var_name="Metrika", value_name="Hodnota")

# 3) Denní součty (všichni uživatelé; bez Transfery)
den_totals = den_user.groupby(["Den"], dropna=False)[["Prijem", "Expedice", "Tonovani", "Celkem_bez_Transfery"]].sum().reset_index()

# 4) Transfery – zvlášť (Den × LocationBucket; počet řádků)
transfer_df = df_f.loc[df_f["is_transfer"]].groupby(["Den", "LocationBucket"], dropna=False).size().reset_index(name="Lines_Transfer")

# -------------------------------
# UI – Tabs
# -------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1) Den × Uživatel (souhrn)", "2) Průřez SKU", "3) Denní součty", "4) Transfery (mimo součty)", "📈 Grafy"
])

with tab1:
    st.subheader("Den × Uživatel – souhrn metrik (bez Transfery)")
    st.caption("Příjem = ST→1, jinak Množství práce; Expedice = ST→1, jinak Množství práce; Tónování = PAL→×24, jinak Množství práce")

    st.dataframe(den_user.sort_values(["Den", "User"]), use_container_width=True, height=420)

with tab2:
    st.subheader("Průřez SKU (Den × Uživatel × SKU) – bez Transfery")
    pivot_on = st.toggle("Pivot (sloupce: Prijem / Expedice / Tonovani)", value=True)
    top_n = st.slider("Top-N SKU (podle součtu metrik bez Transfery v rámci vybraných filtrů)", 5, 100, 20, step=5)

    # Výpočet Top-N v rámci aktuálního filtru: součet přes metriky
    sku_scores = sku_pivot.assign(Soucet=lambda x: x[["Prijem", "Expedice", "Tonovani"]].sum(axis=1))
    # Pokud chceš přesně „pro daný den i uživatele“, můžeš si níže zvolit konkrétní kombo:
    uniq_days = sorted(den_user["Den"].unique())
    uniq_users = sorted(den_user["User"].unique())
    sel_day_for_top = st.selectbox("Den pro Top‑N (volitelné, jinak napříč vybraným rozsahem)", options=["(vše)"] + [str(d) for d in uniq_days], index=0)
    sel_user_for_top = st.selectbox("Uživatel pro Top‑N (volitelné, jinak napříč vybranými)", options=["(všichni)"] + uniq_users, index=0)

    mask_top = pd.Series(True, index=sku_scores.index)
    if sel_day_for_top != "(vše)":
        mask_top &= (sku_scores["Den"].astype(str) == sel_day_for_top)
    if sel_user_for_top != "(všichni)":
        mask_top &= (sku_scores["User"] == sel_user_for_top)

    sku_top = (sku_scores.loc[mask_top]
               .sort_values("Soucet", ascending=False)
               .groupby(["Den", "User", "SKU"], as_index=False)
               .agg({"Prijem":"sum","Expedice":"sum","Tonovani":"sum","Soucet":"sum"}))

    # Vybereme Top-N v rámci každé kombinace Den×User (nebo napříč pokud vybráno "(vše)/(všichni)")
    if sel_day_for_top != "(vše)" or sel_user_for_top != "(všichni)":
        sku_top = sku_top.head(top_n)
    else:
        sku_top = (sku_top
                   .sort_values("Soucet", ascending=False)
                   .head(top_n))

    if pivot_on:
        st.dataframe(sku_pivot.sort_values(["Den", "User", "SKU"]), use_container_width=True, height=360)
    else:
        st.dataframe(sku_long.sort_values(["Den", "User", "SKU", "Metrika"]), use_container_width=True, height=360)

    st.markdown("**Top‑N SKU (viz volby výše)**")
    # Bar chart pro Top‑N – stacked přes metriky
    if not sku_top.empty:
        sku_top_melt = sku_top.melt(id_vars=["Den", "User", "SKU", "Soucet"], value_vars=["Prijem","Expedice","Tonovani"],
                                    var_name="Metrika", value_name="Hodnota")
        chart_top = (
            alt.Chart(sku_top_melt)
            .mark_bar()
            .encode(
                x=alt.X("Hodnota:Q", title="Hodnota"),
                y=alt.Y("SKU:N", sort="-x", title="SKU"),
                color=alt.Color("Metrika:N"),
                tooltip=["Den:N", "User:N", "SKU:N", "Metrika:N", "Hodnota:Q"]
            )
            .properties(height=28 * max(5, len(sku_top["SKU"].unique())), width=800)
        )
        st.altair_chart(chart_top.interactive(), use_container_width=True)
    else:
        st.info("Pro vybrané filtry/top‑N není co zobrazit.")

with tab3:
    st.subheader("Denní součty (všichni uživatelé dohromady, bez Transfery)")
    st.dataframe(den_totals.sort_values("Den"), use_container_width=True, height=380)

with tab4:
    st.subheader("Transfery – Den × Lokace (NEzapočítávat do součtů)")
    st.caption("Lokace jsou sloučeny do 'Adresy (kódy)' pokud vypadají jako regálové adresy (např. F-9-1-1), ostatní texty jsou vykázány zvlášť.")
    if transfer_df.empty:
        st.info("Ve zvolených filtrech nejsou žádné transfery.")
    else:
        st.dataframe(transfer_df.sort_values(["Den", "LocationBucket"]), use_container_width=True, height=360)

with tab5:
    st.subheader("📈 Grafy (respektují vybrané filtry)")
    # Heatmapa Den × Uživatel – Celkem_bez_Transfery
    if not den_user.empty:
        # Bezpečné textové osy
        den_user_plot = den_user.copy()
        den_user_plot["Den_str"] = den_user_plot["Den"].astype(str)
        heat = (
            alt.Chart(den_user_plot)
            .mark_rect()
            .encode(
                x=alt.X("Den_str:N", title="Den"),
                y=alt.Y("User:N", title="Uživatel"),
                color=alt.Color("Celkem_bez_Transfery:Q", title="Celkem (bez Transfery)"),
                tooltip=["Den_str:N", "User:N", "Prijem:Q", "Expedice:Q", "Tonovani:Q", "Celkem_bez_Transfery:Q"]
            )
            .properties(height=420)
        )
        st.altair_chart(heat, use_container_width=True)
    else:
        st.info("Heatmapa: není co zobrazit.")

    st.markdown("---")
    st.markdown("**Stacked bar – struktura metrik po dnech (bez Transfery)**")
    if not den_totals.empty:
        den_totals_plot = den_totals.melt(id_vars=["Den"], value_vars=["Prijem", "Expedice", "Tonovani"],
                                          var_name="Metrika", value_name="Hodnota")
        den_totals_plot["Den_str"] = den_totals_plot["Den"].astype(str)
        stacked = (
            alt.Chart(den_totals_plot)
            .mark_bar()
            .encode(
                x=alt.X("Den_str:N", title="Den"),
                y=alt.Y("Hodnota:Q", title="Hodnota"),
                color=alt.Color("Metrika:N"),
                tooltip=["Den_str:N", "Metrika:N", "Hodnota:Q"]
            )
            .properties(height=320)
        )
        st.altair_chart(stacked, use_container_width=True)
    else:
        st.info("Stacked bar: není co zobrazit.")

    st.markdown("---")
    st.markdown("**Trend – Celkem (bez Transfery) po dnech**")
    if not den_totals.empty:
        trend_df = den_totals.copy()
        trend_df["Den_str"] = trend_df["Den"].astype(str)
        trend = (
            alt.Chart(trend_df)
            .mark_line(point=True)
            .encode(
                x=alt.X("Den_str:N", title="Den"),
                y=alt.Y("Celkem_bez_Transfery:Q", title="Celkem (bez Transfery)"),
                tooltip=["Den_str:N", "Celkem_bez_Transfery:Q"]
            )
            .properties(height=300)
        )
        st.altair_chart(trend.interactive(), use_container_width=True)
    else:
        st.info("Trend: není co zobrazit.")

# -------------------------------
# Export do Excelu
# -------------------------------
def to_excel_bytes(sheets: dict, filename: str = "metriky_denni.xlsx") -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        for name, d in sheets.items():
            # omezíme délku jména listu na 31 znaků
            sheet_name = (name or "Sheet")[:31]
            if isinstance(d, pd.DataFrame):
                d_to_write = d.copy()
            else:
                d_to_write = pd.DataFrame(d)
            d_to_write.to_excel(w, index=False, sheet_name=sheet_name)
    return out.getvalue()

st.sidebar.markdown("---")
st.sidebar.subheader("Export")
export_name = st.sidebar.text_input("Název Excelu", value=f"metriky_denni_{date.today().isoformat()}.xlsx")
if st.sidebar.button("⬇️ Stáhnout Excel"):
    sheets = {
        "den_user": den_user.sort_values(["Den", "User"]),
        "sku_prurez": (sku_pivot.sort_values(["Den", "User", "SKU"])),
        "den_soucty": den_totals.sort_values("Den"),
        "transfery": transfer_df.sort_values(["Den", "LocationBucket"]),
    }
    xls = to_excel_bytes(sheets, filename=export_name)
    st.sidebar.download_button(
        label="Uložit Excel",
        data=xls,
        file_name=export_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# -------------------------------
# Info/validace jednotek
# -------------------------------
# Ukaž varování, pokud existují jednotky mimo ST/PAL
units_other = sorted(set(df_f["Unit_n"].unique()) - {"ST", "PAL"})
if units_other:
    st.warning(
        "V datech se vyskytují i jiné jednotky než **ST/PAL** "
        f"(normalizovaně: {', '.join(units_other)}). "
        "Pro tyto řádky se používá **skutečné množství** bez dalších úprav."
    )

st.success("Hotovo. Všechny souhrny a grafy nezahrnují Transfery. Transfery jsou zobrazeny pouze v samostatné záložce.")