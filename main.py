# africa_manager.py
# ─────────────────────────────────────────────────────────────────────────────
# Provisão & PayRoll – Consolidador África (GUI Tkinter)
# • Executa Provisão OU PayRoll com um clique
# • Log ao vivo (thread-safe)
# • Backup automático antes de mexer nos arquivos principais
# • Pede o usuário do Windows para montar o caminho do PayRoll
# • UI ajustada: sem textos de “Base/Saída” no topo; input do usuário mantido
# ─────────────────────────────────────────────────────────────────────────────

from __future__ import annotations

import os
import re
import shutil
import threading
import queue
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Final, Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from tkinter import Tk, BOTH, LEFT, RIGHT, END, DISABLED, NORMAL, StringVar, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

TZ_FMT_DATE = "%d.%m.%y"
TZ_FMT_TIME = "%H.%M"

BASE_DIR_PROV: Final[Path] = Path(os.environ.get("PROVISIONS_INPUT_DIR", "input/provisions"))
DB_DIR: Final[Path] = Path(os.environ.get("PAYROLL_OUTPUT_DIR", "outputs"))
BACKUP_DIR: Final[Path] = DB_DIR / "Backup Provisão e PayRoll"

ARQ_SAIDA_PROV: Final[Path] = DB_DIR / "Provisions and Reversals.xlsx"
ARQ_SAIDA_PAY: Final[Path] = DB_DIR / "Payroll_Africa.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# PROVISÃO – constantes do processo
# ─────────────────────────────────────────────────────────────────────────────
MIN_YEAR: Final[int] = 2025
SHEET_DADOS_PROV: Final[str] = "África - Provisão x Estorno"
SHEET_DEPARA_PROV: Final[str] = "Depara"
TABLE_NAME_PROV: Final[str] = "africa_provisao"
COLS_SAIDA_PROV = ["OPERACAO", "COMP", "VALOR", "OBS"]
COLS_TOTAIS_PROV = COLS_SAIDA_PROV + ["Un_Neg", "Un_Neg.Nomenclatura"]
MONTH_MAP_PROV: Final[dict[str, str]] = {
    "JANUARY": "01",
    "FEBRUARY": "02",
    "MARCH": "03",
    "APRIL": "04",
    "MAY": "05",
    "JUNE": "06",
    "JULY": "07",
    "AUGUST": "08",
    "SEPTEMBER": "09",
    "OCTOBER": "10",
    "NOVEMBER": "11",
    "DECEMBER": "12",
}
REGEX_ARQ_PROV = re.compile(
    r"^PROVISÃO\s+RECEITAS\s+ÁFRICA\s+\d{4}\s*-\s*DEMO_ORGANIZATION\s+LOGISTICS\.xlsx$",
    re.IGNORECASE,
)
REGEX_SHEET_PROV = re.compile(r"^(?P<mes>[A-Z]+)\s+PROVISION(?:\s+(?P<ano>\d{4}))?$", re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────────────────
# PAYROLL – constantes do processo
# ─────────────────────────────────────────────────────────────────────────────
ANO_INICIAL_PAY = 2025
PRIMEIRO_MES_ANO_INICIAL_PAY = 5  # mantido para preservar histórico/comentários

MONTH_MAP_PAY = {
    "JANUARY": "01",
    "FEBRUARY": "02",
    "MARCH": "03",
    "APRIL": "04",
    "MAY": "05",
    "JUNE": "06",
    "JULY": "07",
    "AUGUST": "08",
    "SEPTEMBER": "09",
    "OCTOBER": "10",
    "NOVEMBER": "11",
    "DECEMBER": "12",
}

COLUNAS_FINAIS_PAY = [
    "Region Description",
    "Employee Last Name",
    "Employee First Name",
    "Position",
    "Comp",
    "EMP",
    "13th Bonus",
    "Additional Functions",
    "Arbitration Award (Loss of Income)",
    "Backdated Overtime @ 1.5",
    "Backdated Overtime @ Double Time",
    "Backpay - Regular Salary",
    "Backpay Subsistence Allowance",
    "Basic Pay",
    "Provisão 13º",
    "Cost of living allowance",
    "Housing Allowance",
    "Meals",
    "Payment of employee debt (gross up)",
    "PLR Bonus",
    "Reimbursment of BR Tax",
    "School Allowance",
    "Seniority",
    "Transfer Allowance",
    "Utilities",
    "Beer Allowance",
    "Leave Pay",
    "Medical Aid Allowance",
    "Night Allowance",
    "Other Lump Sums",
    "Overtime @ 1.5",
    "Overtime @ Double Time",
    "Productive incentive bonus",
    "Shift All 20%",
    "Shift All 30%",
    "Shift Allowance Back pay 30%",
    "Subsistence Allowance",
    "Travel Allowance",
    "Unpaid Hours",
    "Unpaid Time",
    "Bargaining Council Levy ER",
    "ER Holiday Contribution",
    "ER Leave Contribution",
    "ER Sick Contribution",
    "ER Wellness Contribution",
    "Med Aid ER",
    "Provident ER Contribution",
    "Skills Development Levy",
    "UIF Adjust Employer",
    "UIF Employer",
]

# Mapeamento de aliases / nomenclaturas parecidas
# Mantém SEMPRE o nome canônico em COLUNAS_FINAIS_PAY
COLUNAS_ALTERNATIVAS_PAY: dict[str, list[str]] = {
    "Region Description": [
        "Org. Unit",
        "Org Unit",
        "Org.Unit",
        "Organizational Unit",
        "Organisation Unit",
        "Region",
        "Region Desc",
    ],
    "Additional Functions": [
        "Add Function",
        "Add Functions",
        "Add. Function",
        "Add. Functions",
        "Additional Function",
    ],
    "13th Bonus": [
        "13th bonus",
        "13th_bonus",
        "13th Bonus ",
        "Thirteenth Bonus",
    ],
    "Reimbursment of BR Tax": [
        "Reimbursement of BR Tax",
        "Reimbursment BR Tax",
        "Reimbursement BR Tax",
    ],
    "Cost of living allowance": [
        "Cost of Living Allowance",
        "Cost of living",
    ],
    "Housing Allowance": [
        "Housing allowance",
    ],
    "Payment of employee debt (gross up)": [
        "Payment employee debt (gross up)",
        "Payment of employee debt gross up",
    ],
}

REGEX_EMP_PAY = re.compile(r"(NWL|NEPO|EXPATS)", flags=re.IGNORECASE)

KEYWORDS_PAY = (
    "payroll recon",
    "payroll reconciliation",
    "payroll reonciliation",
    "payroll report",
)

# Posições que NÃO devem gerar Provisão 13º
POSICOES_SEM_PROVISAO_13 = {
    "DRIVER",
    "DRIVER ASSISTENT",
    "DRIVER ASSISTANT",
    "DRIVER ASSIST",
    "DRIVER ASSISTENTE",
    "DRIVER HELPER",
    "DRIVER - HELPER",
    "DRIVER-HELPER",
}

# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────
def ts_prefix(dt: datetime | None = None) -> str:
    dt = dt or datetime.now()
    return f"{dt.strftime(TZ_FMT_DATE)} - {dt.strftime(TZ_FMT_TIME)}"


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def backup_file(target: Path, log: Callable[[str], None]) -> None:
    """Faz backup do arquivo `target` em BACKUP_DIR com prefixo 'DD.MM.AA - HH.MM - '."""
    if not target.exists():
        log(f"🗃️  Backup: arquivo ainda não existe → pulando ({target.name}).")
        return

    ensure_dir(BACKUP_DIR)
    dst_name = f"{ts_prefix()} - {target.name}"
    dst_path = BACKUP_DIR / dst_name

    try:
        shutil.copy2(str(target), str(dst_path))
        log(f"🗄️  Backup criado: {dst_path}")
    except Exception as e:
        log(f"⚠️  Falha ao criar backup de {target.name}: {e}")


def text_to_float(txt) -> float:
    if pd.isna(txt):
        return 0.0
    if isinstance(txt, (int, float)):
        return float(txt)

    s = str(txt).replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def prox_comp(comp: str) -> str:
    dt = datetime.strptime(comp, "%Y%m")
    ano = dt.year + (dt.month == 12)
    mes = 1 if dt.month == 12 else dt.month + 1
    return f"{ano}{mes:02d}"


def _normalizar_texto_pay(txt: str) -> str:
    """
    Normaliza texto para comparação:
    - maiúsculas
    - remove acentos
    - troca underscore por espaço
    - remove pontuação
    - colapsa espaços
    """
    s = str(txt).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("_", " ")
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _normalizar_nome_coluna_pay(col: str) -> str:
    return _normalizar_texto_pay(col)


def _serie_base_vazia(df: pd.DataFrame) -> pd.Series:
    return pd.Series([pd.NA] * len(df), index=df.index, dtype="object")


def _harmonizar_colunas_pay(df: pd.DataFrame, log: Callable[[str], None]) -> pd.DataFrame:
    """
    Padroniza nomes e combina colunas equivalentes/parecidas
    mantendo os nomes finais definidos em COLUNAS_FINAIS_PAY.
    """
    if df.empty:
        return df.reindex(columns=COLUNAS_FINAIS_PAY, fill_value=pd.NA)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    mapa_norm_para_cols: dict[str, list[str]] = {}
    for col in df.columns:
        mapa_norm_para_cols.setdefault(_normalizar_nome_coluna_pay(col), []).append(col)

    for col_destino in COLUNAS_FINAIS_PAY:
        candidatos_nomes = [col_destino, *COLUNAS_ALTERNATIVAS_PAY.get(col_destino, [])]

        cols_encontradas: list[str] = []
        for nome in candidatos_nomes:
            norm = _normalizar_nome_coluna_pay(nome)
            for col_real in mapa_norm_para_cols.get(norm, []):
                if col_real not in cols_encontradas:
                    cols_encontradas.append(col_real)

        if col_destino in df.columns and col_destino not in cols_encontradas:
            cols_encontradas.insert(0, col_destino)

        if cols_encontradas:
            serie_resultado = _serie_base_vazia(df)
            for col_real in cols_encontradas:
                serie_resultado = serie_resultado.combine_first(df[col_real])

            df[col_destino] = serie_resultado

            cols_alias_usadas = [c for c in cols_encontradas if c != col_destino]
            if cols_alias_usadas:
                log(f"     🔁 Coluna(s) harmonizada(s) em '{col_destino}': {', '.join(cols_alias_usadas)}")
        else:
            df[col_destino] = pd.NA

    return df[COLUNAS_FINAIS_PAY]


def _emp_from_filename(nome_arquivo: str) -> str:
    fname_up = nome_arquivo.upper()

    match = REGEX_EMP_PAY.search(fname_up)
    if match:
        emp_val = match.group(1).upper()
    elif "NEPO" in fname_up or re.search(r"\bNEP\b", fname_up):
        emp_val = "NEPO"
    elif "NWL" in fname_up:
        emp_val = "NWL"
    else:
        emp_val = "EXPATS"

    if re.search(r"\bEXPAT\S*\b", fname_up):
        return "EXPATS"
    if re.search(r"\bNWL\b", fname_up):
        return "NWL"
    if re.search(r"\bNEPO\b", fname_up) or re.search(r"\bNEP\b", fname_up):
        return "NEPO"

    return emp_val


def _deve_calcular_provisao_13(position) -> bool:
    pos_norm = _normalizar_texto_pay(position)
    if not pos_norm:
        return True
    return pos_norm not in POSICOES_SEM_PROVISAO_13


def _calcular_provisao_13(df: pd.DataFrame) -> pd.Series:
    """
    Calcula a Provisão 13º = Basic Pay / 12
    somente quando o cargo NÃO for Driver nem Driver - Assistent.
    """
    basic_pay = pd.to_numeric(df["Basic Pay"], errors="coerce").fillna(0)
    posicoes_validas = df["Position"].apply(_deve_calcular_provisao_13)
    return basic_pay.where((basic_pay != 0) & posicoes_validas, 0) / 12


# ─────────────────────────────────────────────────────────────────────────────
# PROVISÃO – núcleo
# ─────────────────────────────────────────────────────────────────────────────
def ler_consolidado_prov(log: Callable[[str], None]) -> tuple[pd.DataFrame, set[str], pd.DataFrame]:
    """Retorna df_exist, comps_com_provisao e depara."""
    if not ARQ_SAIDA_PROV.exists():
        log("ℹ️  Consolidado ainda não existe; será criado vazio.")
        return (
            pd.DataFrame(columns=COLS_TOTAIS_PROV),
            set(),
            pd.DataFrame(columns=["Desc_BDO", "Un_Neg", "Descrição"]),
        )

    df_exist = pd.read_excel(ARQ_SAIDA_PROV, sheet_name=SHEET_DADOS_PROV, dtype={"COMP": str})
    comps_prov = set(df_exist.loc[df_exist["OBS"] == "PROVISÃO", "COMP"])

    try:
        depara = pd.read_excel(ARQ_SAIDA_PROV, sheet_name=SHEET_DEPARA_PROV).rename(
            columns=str.strip,
            errors="ignore",
        )
    except Exception:
        depara = pd.DataFrame(columns=["Desc_BDO", "Un_Neg", "Descrição"])

    return df_exist, comps_prov, depara


def localizar_sheet_recente_prov(log: Callable[[str], None]) -> tuple[Path, str, str] | None:
    """
    Percorre todas as pastas/arquivos e devolve (arquivo, nome_sheet, comp) da
    competência mais nova encontrada. Retorna None se nada for encontrado.
    """
    candidato: tuple[int, Path, str] | None = None

    for ano_dir in sorted(p for p in BASE_DIR_PROV.iterdir() if p.is_dir() and p.name.isdigit()):
        if int(ano_dir.name) < MIN_YEAR:
            continue

        alvo = next((f for f in ano_dir.iterdir() if f.is_file() and REGEX_ARQ_PROV.match(f.name)), None)
        if not alvo:
            continue

        xls = pd.ExcelFile(alvo, engine="openpyxl")
        for sheet in xls.sheet_names:
            m = REGEX_SHEET_PROV.match(sheet)
            if not m:
                continue

            mes = m.group("mes").upper()
            ano_sheet = int(m.group("ano") or ano_dir.name)
            comp = int(f"{ano_sheet}{MONTH_MAP_PROV[mes]}")

            if (candidato is None) or (comp > candidato[0]):
                candidato = (comp, alvo, sheet)

    if not candidato:
        return None

    comp_str = str(candidato[0])
    return candidato[1], candidato[2], comp_str


def extrair_linhas_prov(arquivo: Path, sheet: str, comp: str, log: Callable[[str], None]) -> list[dict]:
    dados: list[dict] = []
    xls_df = pd.read_excel(arquivo, sheet_name=sheet, header=None, skiprows=1)

    xls_df.columns = xls_df.iloc[0]
    xls_df = xls_df.drop(index=0).reset_index(drop=True)
    xls_df.rename(columns=lambda c: str(c).strip().title(), inplace=True)

    if {"Operação", "Valor Líquido"} - set(xls_df.columns):
        log(f"⚠️  Sheet '{sheet}' sem colunas essenciais (Operação/Valor Líquido).")
        return dados

    xls_df = xls_df[~xls_df["Operação"].astype(str).str.upper().str.startswith(("TOTAL", "OBS:"))]
    xls_df = xls_df[xls_df["Operação"].notna() & (xls_df["Operação"].astype(str).str.strip() != "")]

    for _, row in xls_df.iterrows():
        oper = str(row["Operação"]).strip()
        valor = text_to_float(row["Valor Líquido"])

        dados.append(dict(OPERACAO=oper, COMP=comp, VALOR=-valor, OBS="PROVISÃO"))
        prox = prox_comp(comp)
        if not (oper == "Alrode Brewery (T1)" and prox == "202504"):
            dados.append(dict(OPERACAO=oper, COMP=prox, VALOR=valor, OBS="ESTORNO"))

    return dados


def executar_provisao(log: Callable[[str], None]) -> None:
    log("🔎 Lendo consolidado de Provisões…")
    df_exist, comps_prov, depara = ler_consolidado_prov(log)

    log("🔎 Buscando sheet mais recente nos arquivos-fonte…")
    info = localizar_sheet_recente_prov(log)
    if not info:
        log("❗ Nenhum arquivo/sheet de provisão encontrado.")
        return

    arq_src, sheet_src, comp_src = info
    log(f"📄 Encontrado: {arq_src.name} › {sheet_src} (COMP {comp_src})")

    if comp_src in comps_prov:
        log(f"✅ Já existe PROVISÃO para {comp_src} — nada a fazer.")
        return

    novos_registros = extrair_linhas_prov(arq_src, sheet_src, comp_src, log)
    if not novos_registros:
        log("ℹ️  Nenhum dado extraído — encerrando.")
        return

    df_novos = pd.DataFrame(novos_registros, columns=COLS_SAIDA_PROV)

    if not depara.empty:
        df_novos = (
            df_novos
            .merge(
                depara[["Desc_BDO", "Un_Neg", "Descrição"]],
                left_on="OPERACAO",
                right_on="Desc_BDO",
                how="left",
            )
            .drop(columns="Desc_BDO")
            .rename(columns={"Descrição": "Un_Neg.Nomenclatura"})
        )

    df_novos = df_novos.reindex(columns=COLS_TOTAIS_PROV, fill_value=pd.NA)

    df_total = pd.concat([df_exist, df_novos], ignore_index=True).drop_duplicates(
        subset=COLS_SAIDA_PROV,
        keep="first",
    )

    backup_file(ARQ_SAIDA_PROV, log)

    if ARQ_SAIDA_PROV.exists():
        with pd.ExcelWriter(
            ARQ_SAIDA_PROV,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as wrt:
            df_total.to_excel(wrt, index=False, sheet_name=SHEET_DADOS_PROV)
    else:
        with pd.ExcelWriter(ARQ_SAIDA_PROV, engine="openpyxl", mode="w") as wrt:
            df_total.to_excel(wrt, index=False, sheet_name=SHEET_DADOS_PROV)
            depara.to_excel(wrt, index=False, sheet_name=SHEET_DEPARA_PROV)

    wb = load_workbook(ARQ_SAIDA_PROV)
    ws = wb[SHEET_DADOS_PROV]
    ref = f"A1:{get_column_letter(df_total.shape[1])}{df_total.shape[0] + 1}"

    if TABLE_NAME_PROV in ws.tables:
        del ws.tables[TABLE_NAME_PROV]

    tbl = Table(displayName=TABLE_NAME_PROV, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tbl)
    wb.save(ARQ_SAIDA_PROV)
    wb.close()

    log(f"✅ Provisão concluída. Linhas novas: {len(df_novos)} → {ARQ_SAIDA_PROV}")


# ─────────────────────────────────────────────────────────────────────────────
# PAYROLL – núcleo
# ─────────────────────────────────────────────────────────────────────────────
def montar_base_payroll(usuario_windows: str) -> Path:
    return Path(os.environ.get("PAYROLL_BASE_DIR", "input/payroll"))


def executar_payroll(usuario_windows: str, log: Callable[[str], None]) -> None:
    base_root_dir = montar_base_payroll(usuario_windows)
    log(f"📂 Base PayRoll: {base_root_dir}")

    if not base_root_dir.is_dir():
        raise FileNotFoundError(f"Diretório não encontrado: {base_root_dir}")

    if ARQ_SAIDA_PAY.exists():
        df_final = pd.read_excel(ARQ_SAIDA_PAY, engine="openpyxl")
        df_final = _harmonizar_colunas_pay(df_final, log)
        df_final["Provisão 13º"] = _calcular_provisao_13(df_final)
        log(f"ℹ️  Saída existente carregada ({len(df_final)} linhas).")
    else:
        df_final = pd.DataFrame(columns=COLUNAS_FINAIS_PAY)
        log("ℹ️  Criando novo consolidado PayRoll.")

    # controle do que já existe por COMP + EMP
    existing: dict[str, set[str]] = {}
    for _, r in df_final[["Comp", "EMP"]].dropna().iterrows():
        existing.setdefault(str(r["Comp"]), set()).add(str(r["EMP"]))

    for ano_dir in sorted(p for p in base_root_dir.iterdir() if p.is_dir()):
        ano_str = ano_dir.name
        if not ano_str.isdigit():
            continue

        ano_int = int(ano_str)
        if ano_int < ANO_INICIAL_PAY:
            continue

        for pasta in sorted(p for p in ano_dir.iterdir() if p.is_dir()):
            folder_upper = pasta.name.upper()

            month_key = next((m for m in MONTH_MAP_PAY if m in folder_upper), None)
            if month_key is None:
                continue

            comp_val_pasta = f"{ano_str}{MONTH_MAP_PAY[month_key]}"

            arquivos = []
            for f in pasta.iterdir():
                if f.suffix.lower() != ".xlsx":
                    continue
                fname_key = f.name.lower().replace("_", " ")
                if any(kw in fname_key for kw in KEYWORDS_PAY):
                    arquivos.append(f)

            if not arquivos:
                continue

            log(f"🔎 {ano_str} / {pasta.name} — {len(arquivos)} arquivo(s)")

            for fpath in arquivos:
                comp_val = comp_val_pasta
                emp_val = _emp_from_filename(fpath.name)

                if comp_val in existing and emp_val in existing[comp_val]:
                    log(f"  ↷  Skip {comp_val}/{emp_val} já consolidado.")
                    continue

                log(f"  ➕ Lendo: {fpath.name}  →  COMP={comp_val}, EMP={emp_val}")
                try:
                    df = pd.read_excel(fpath, skiprows=10, engine="openpyxl")
                except Exception as exc:
                    log(f"     ❌ Erro ao ler '{fpath.name}': {exc}")
                    continue

                df = _harmonizar_colunas_pay(df, log)
                df["Comp"] = comp_val
                df["EMP"] = emp_val

                antes = len(df)
                df.dropna(
                    subset=[
                        "Region Description",
                        "Employee Last Name",
                        "Employee First Name",
                        "Position",
                    ],
                    how="all",
                    inplace=True,
                )

                # Regra expats
                if emp_val == "EXPATS":
                    df["Region Description"] = df["Region Description"].where(
                        df["Region Description"].notna()
                        & (df["Region Description"].astype(str).str.strip() != ""),
                        "Head Office",
                    )

                # Regra da provisão 13º
                df["Provisão 13º"] = _calcular_provisao_13(df)

                log(f"     Linhas lidas: {antes}  →  mantidas: {len(df)}")
                bloqueadas = (~df["Position"].apply(_deve_calcular_provisao_13)).sum()
                if bloqueadas:
                    log(f"     🚫 Provisão 13º zerada para {bloqueadas} linha(s) com cargo Driver/Driver - Assistent.")

                df_final = pd.concat([df_final, df], ignore_index=True)
                existing.setdefault(comp_val, set()).add(emp_val)

    # ajuste final para expats
    expats_mask = df_final["EMP"] == "EXPATS"
    df_final.loc[expats_mask, "Region Description"] = df_final.loc[expats_mask, "Region Description"].where(
        df_final.loc[expats_mask, "Region Description"].notna()
        & (df_final.loc[expats_mask, "Region Description"].astype(str).str.strip() != ""),
        "Head Office",
    )

    # harmoniza novamente para garantir layout final
    df_final = _harmonizar_colunas_pay(df_final, log)

    # recalcula sempre no consolidado inteiro, aplicando a regra nova
    df_final["Provisão 13º"] = _calcular_provisao_13(df_final)

    backup_file(ARQ_SAIDA_PAY, log)

    df_final.to_excel(ARQ_SAIDA_PAY, index=False, engine="openpyxl")

    wb = load_workbook(ARQ_SAIDA_PAY)
    ws = wb.active
    last_col = get_column_letter(df_final.shape[1])
    last_row = df_final.shape[0] + 1
    ref = f"A1:{last_col}{last_row}"

    if "PayRollDBO" in ws.tables:
        del ws.tables["PayRollDBO"]

    tbl = Table(displayName="PayRollDBO", ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tbl)
    wb.save(ARQ_SAIDA_PAY)
    wb.close()

    log(f"✅ PayRoll concluído. Linhas totais: {len(df_final)} → {ARQ_SAIDA_PAY}")


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("África – Provisão & PayRoll")
        W, H = 980, 640
        self.geometry(f"{W}x{H}+{(self.winfo_screenwidth()-W)//2}+{(self.winfo_screenheight()-H)//2}")
        self.minsize(860, 560)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TProgressbar", thickness=10)

        try:
            user_name = os.getlogin()
        except Exception:
            user_name = ""

        self.usuario_windows = StringVar(value=user_name)
        self.status = StringVar(value="Pronto")
        self._running = False

        self.log_q: queue.Queue[str] = queue.Queue()
        self.done_q: queue.Queue[tuple[bool, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_ui()
        self.after(100, self._poll_queues)

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=14, pady=(12, 6))
        ttk.Label(top, text="África – Provisão & PayRoll", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            top,
            text="Selecione um processo. Backups são criados automaticamente antes de gravar.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        actions = ttk.Frame(self)
        actions.pack(fill="x", padx=14, pady=8)

        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)

        prov = ttk.LabelFrame(actions, text="Provisão")
        prov.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=0)
        self.btn_prov = ttk.Button(prov, text="▶ Executar Provisão", command=self._start_provisao, width=24)
        self.btn_prov.pack(padx=8, pady=(10, 10), anchor="w")

        pay = ttk.LabelFrame(actions, text="PayRoll")
        pay.grid(row=0, column=1, sticky="nsew", padx=6, pady=0)
        self.btn_pay = ttk.Button(pay, text="▶ Executar PayRoll", command=self._start_payroll, width=24)
        self.btn_pay.pack(padx=8, pady=(10, 10), anchor="w")

        user = ttk.LabelFrame(actions, text="Usuário do Windows")
        user.grid(row=0, column=2, sticky="nsew", padx=(6, 0), pady=0)

        user.columnconfigure(1, weight=1)
        ttk.Label(user, text="Usuário:").grid(row=0, column=0, padx=(8, 4), pady=(10, 10), sticky="w")
        self.ent_user = ttk.Entry(user, textvariable=self.usuario_windows, width=28)
        self.ent_user.grid(row=0, column=1, padx=(0, 8), pady=(10, 10), sticky="we")

        prog_fr = ttk.Frame(self)
        prog_fr.pack(fill="x", padx=14, pady=(0, 8))
        self.pbar = ttk.Progressbar(prog_fr, orient="horizontal", mode="indeterminate")
        self.pbar.pack(fill="x", expand=True)
        ttk.Label(prog_fr, textvariable=self.status).pack(anchor="w", pady=(6, 0))

        log_fr = ttk.Frame(self)
        log_fr.pack(fill=BOTH, expand=True, padx=14, pady=(4, 8))
        self.logbox = ScrolledText(log_fr, height=18, font=("Consolas", 10))
        self.logbox.pack(fill=BOTH, expand=True)
        self._log("🟢 Pronto.\n")

        btn_fr = ttk.Frame(self)
        btn_fr.pack(fill="x", padx=14, pady=(4, 12))
        ttk.Button(btn_fr, text="Abrir pasta de Backup", command=self._open_backup).pack(side=LEFT)
        ttk.Button(btn_fr, text="Abrir pasta do DB", command=lambda: self._open_folder(DB_DIR)).pack(side=LEFT, padx=(6, 0))
        ttk.Button(btn_fr, text="Fechar", command=self.destroy).pack(side=RIGHT)

    def _log(self, text: str):
        self.logbox.configure(state=NORMAL)
        self.logbox.insert(END, text if text.endswith("\n") else text + "\n")
        self.logbox.see(END)
        self.logbox.configure(state=DISABLED)

    def _set_running(self, running: bool, label: str):
        self._running = running
        self.btn_prov.configure(state=DISABLED if running else NORMAL)
        self.btn_pay.configure(state=DISABLED if running else NORMAL)
        self.ent_user.configure(state=DISABLED if running else NORMAL)
        self.status.set(label)
        if running:
            self.pbar.start(12)
        else:
            self.pbar.stop()

    def _open_folder(self, p: Path):
        try:
            os.startfile(str(p))
        except Exception:
            messagebox.showinfo("Info", f"Não foi possível abrir:\n{p}")

    def _open_backup(self):
        ensure_dir(BACKUP_DIR)
        self._open_folder(BACKUP_DIR)

    def _start_provisao(self):
        if self._running:
            return
        self._set_running(True, "Executando Provisão…")
        self._log("🚀 Iniciando Provisão…")
        self._kick_worker(target="provisao")

    def _start_payroll(self):
        if self._running:
            return

        user = self.usuario_windows.get().strip()
        if not user:
            messagebox.showwarning("Atenção", "Informe o usuário do Windows para executar o PayRoll.")
            return

        self._set_running(True, "Executando PayRoll…")
        self._log(f"🚀 Iniciando PayRoll para usuário '{user}'…")
        self._kick_worker(target="payroll", user=user)

    def _kick_worker(self, target: str, user: str | None = None):
        for q in (self.log_q, self.done_q):
            while not q.empty():
                try:
                    q.get_nowait()
                except queue.Empty:
                    break

        def run():
            def log_fn(msg: str):
                self.log_q.put(msg)

            try:
                if target == "provisao":
                    executar_provisao(log_fn)
                else:
                    executar_payroll(user or "", log_fn)
                self.done_q.put((True, "Concluído."))
            except Exception:
                self.done_q.put((False, traceback.format_exc()))

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def _poll_queues(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self._log(msg)
        except queue.Empty:
            pass

        try:
            ok, msg = self.done_q.get_nowait()
            self._set_running(False, "Concluído" if ok else "Falhou")
            self._log(("🏁 " if ok else "❗ ") + (msg.strip() or ""))
        except queue.Empty:
            pass

        self.after(120, self._poll_queues)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    ensure_dir(DB_DIR)
    ensure_dir(BACKUP_DIR)
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()