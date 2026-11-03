#!/usr/bin/env python3
"""Consolida várias planilhas de despesas em um relatório Excel validado."""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

ROOT = Path(__file__).resolve().parents[1]
ENTRADA = ROOT / "dados_entrada"
OUT = ROOT / "output"
LOG_DIR = ROOT / "logs"

COLUNAS_ESPERADAS = [
    "Ano",
    "Mês",
    "Poder",
    "Código órgão",
    "Órgão setorial",
    "Código da unidade",
    "Unidade gestora",
    "Código da despesa",
    "Categoria da despesa",
    "Código do grupo de despesa",
    "Grupo de despesa",
    "Código da modalidade",
    "Modalidade",
    "Dotação atualizada",
    "Dotação inicial",
    "Empenhado",
    "Liquidado",
    "Pago",
]

NUM_COLS = ["Dotação atualizada", "Dotação inicial", "Empenhado", "Liquidado", "Pago"]


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_DIR / f"consolidar_{ts}.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger(__name__)


def parse_br_number(val) -> float | None:
    """Aceita 1.234.567,89 (BR), 1234567.89 (US) ou 1234567,89."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().strip('"')
    if s == "" or s.upper() in {"N/A", "NA", "NULL", "-", "#N/A", "NAN"}:
        return None
    # remove currency symbols / spaces
    for ch in ("R$", " "):
        s = s.replace(ch, "")
    if "," in s and "." in s:
        # BR: milhar com ponto, decimal com vírgula
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        # só vírgula → decimal BR
        s = s.replace(",", ".")
    # só ponto → já está em formato decimal US
    try:
        return float(s)
    except ValueError:
        return None


def carregar_arquivos(log: logging.Logger) -> tuple[pd.DataFrame, pd.DataFrame]:
    files = sorted(ENTRADA.glob("despesas_*.xlsx"))
    if not files:
        raise FileNotFoundError(
            f"Nenhum arquivo em {ENTRADA}. Rode primeiro: python src/preparar_entradas.py"
        )
    frames = []
    rejeitados = []
    for f in files:
        df = pd.read_excel(f, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        df["_arquivo_origem"] = f.name
        # valida colunas mínimas
        faltando = [c for c in ["Ano", "Mês", "Empenhado", "Liquidado", "Pago"] if c not in df.columns]
        if faltando:
            rej = df.copy()
            rej["motivo"] = f"colunas ausentes: {faltando}"
            rejeitados.append(rej)
            log.warning("%s: colunas ausentes %s — arquivo rejeitado", f.name, faltando)
            continue

        for col in NUM_COLS:
            if col not in df.columns:
                df[col] = None
            parsed = df[col].map(parse_br_number)
            invalid_mask = df[col].notna() & (df[col].astype(str).str.strip() != "") & parsed.isna()
            # também rejeita se todos os 3 valores principais forem nulos após parse
            df[col + "_num"] = parsed

        # linha válida se Empenhado/Liquidado/Pago puderem ser interpretados quando preenchidos
        problemas = []
        for col in ["Empenhado", "Liquidado", "Pago"]:
            raw = df[col].fillna("").astype(str).str.strip()
            parsed = df[col + "_num"]
            bad = (raw != "") & parsed.isna()
            problemas.append(bad)
        bad_any = problemas[0]
        for b in problemas[1:]:
            bad_any = bad_any | b

        if bad_any.any():
            rej = df.loc[bad_any].copy()
            rej["motivo"] = "valor monetário inválido"
            rejeitados.append(rej)
            log.info("%s: %d linhas rejeitadas por valor inválido", f.name, bad_any.sum())

        ok = df.loc[~bad_any].copy()
        for col in NUM_COLS:
            ok[col] = ok[col + "_num"]
        frames.append(ok)
        log.info("%s: %d linhas aceitas", f.name, len(ok))

    valid = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    rej_df = pd.concat(rejeitados, ignore_index=True) if rejeitados else pd.DataFrame()
    return valid, rej_df


def montar_resumo(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    base = df.copy()
    base["Ano"] = base["Ano"].astype(str).str.strip()
    base["Mês"] = base["Mês"].astype(str).str.strip().str.zfill(2)
    base["Órgão setorial"] = base["Órgão setorial"].astype(str).str.strip()
    base["Grupo de despesa"] = base.get("Grupo de despesa", pd.Series(index=base.index)).astype(str).str.strip()

    por_mes = (
        base.groupby(["Ano", "Mês"], as_index=False)
        .agg(
            empenhado=("Empenhado", "sum"),
            liquidado=("Liquidado", "sum"),
            pago=("Pago", "sum"),
            linhas=("_arquivo_origem", "count"),
        )
        .sort_values(["Ano", "Mês"])
    )
    por_orgao = (
        base.groupby(["Ano", "Órgão setorial"], as_index=False)
        .agg(empenhado=("Empenhado", "sum"), liquidado=("Liquidado", "sum"), pago=("Pago", "sum"))
        .sort_values("pago", ascending=False)
    )
    por_grupo = (
        base.groupby(["Ano", "Grupo de despesa"], as_index=False)
        .agg(empenhado=("Empenhado", "sum"), liquidado=("Liquidado", "sum"), pago=("Pago", "sum"))
        .sort_values("pago", ascending=False)
    )
    kpis = pd.DataFrame(
        [
            {
                "metricas": "total_linhas_validas",
                "valor": len(base),
            },
            {"metricas": "total_empenhado", "valor": float(base["Empenhado"].fillna(0).sum())},
            {"metricas": "total_liquidado", "valor": float(base["Liquidado"].fillna(0).sum())},
            {"metricas": "total_pago", "valor": float(base["Pago"].fillna(0).sum())},
            {
                "metricas": "pct_pago_sobre_empenhado",
                "valor": float(
                    base["Pago"].fillna(0).sum() / base["Empenhado"].fillna(0).sum()
                    if base["Empenhado"].fillna(0).sum()
                    else 0
                ),
            },
            {"metricas": "orgaos_distintos", "valor": base["Órgão setorial"].nunique()},
            {"metricas": "arquivos_origem", "valor": base["_arquivo_origem"].nunique()},
        ]
    )
    detalhe_cols = [
        c
        for c in [
            "Ano",
            "Mês",
            "Órgão setorial",
            "Unidade gestora",
            "Grupo de despesa",
            "Modalidade",
            "Empenhado",
            "Liquidado",
            "Pago",
            "_arquivo_origem",
        ]
        if c in base.columns
    ]
    return {
        "KPIs": kpis,
        "Por_Mes": por_mes,
        "Por_Orgao": por_orgao.head(50),
        "Por_Grupo": por_grupo,
        "Detalhe": base[detalhe_cols],
    }


def estilo_cabecalho(ws):
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def escrever_excel(abas: dict[str, pd.DataFrame], rejeitados: pd.DataFrame, path: Path):
    wb = Workbook()
    # remove default
    default = wb.active
    wb.remove(default)

    for nome, df in abas.items():
        ws = wb.create_sheet(nome[:31])
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        if ws.max_row >= 1 and ws.max_column >= 1:
            estilo_cabecalho(ws)
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col[:50]:
                max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)
        # formata números nas colunas monetárias
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, (int, float)) and cell.value is not None:
                    hdr = ws.cell(1, cell.column).value
                    if hdr and any(
                        k in str(hdr).lower()
                        for k in ["empenhado", "liquidado", "pago", "dotação", "valor", "total"]
                    ):
                        cell.number_format = "#,##0.00"

    ws = wb.create_sheet("Rejeitados")
    if rejeitados.empty:
        ws.append(["motivo", "obs"])
        ws.append(["(nenhuma linha rejeitada)", ""])
    else:
        cols = [c for c in rejeitados.columns if not str(c).endswith("_num")]
        for r in dataframe_to_rows(rejeitados[cols], index=False, header=True):
            ws.append(r)
        estilo_cabecalho(ws)
    wb.save(path)


def main() -> int:
    log = setup_logging()
    valid, rejeitados = carregar_arquivos(log)
    log.info("Total válidos=%d rejeitados=%d", len(valid), len(rejeitados))
    if not rejeitados.empty:
        rej_path = OUT / "rejeitados.csv"
        rejeitados.to_csv(rej_path, index=False, encoding="utf-8")
        log.info("Log de rejeitados: %s", rej_path)

    abas = montar_resumo(valid)
    out_xlsx = OUT / "relatorio_despesas_consolidado.xlsx"
    escrever_excel(abas, rejeitados, out_xlsx)
    log.info("Relatório gerado: %s", out_xlsx)

    # também um CSV resumo
    abas["Por_Mes"].to_csv(OUT / "resumo_por_mes.csv", index=False, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
