#!/usr/bin/env python3
"""Baixa CSV público (Prefeitura do Recife) e divide em várias planilhas de entrada."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
ENTRADA = ROOT / "dados_entrada"
LOG_DIR = ROOT / "logs"

# Despesas por órgão — Portal de Dados Abertos da Prefeitura do Recife
URLS = {
    "2024": "https://dados.recife.pe.gov.br/dataset/63f7941b-9a1e-4a49-893b-d5f16e2134d8/resource/32880df9-87da-46fa-b77b-7fbfbf75aeaf/download/despesas-orgao-2024.csv",
    "2025": "https://dados.recife.pe.gov.br/dataset/63f7941b-9a1e-4a49-893b-d5f16e2134d8/resource/349f8b0a-8d51-4905-b3b3-eca817911b3e/download/despesas-orgao-2025.csv",
}


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_DIR / "preparar_entradas.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger(__name__)


def baixar(url: str) -> pd.DataFrame:
    r = requests.get(url, timeout=120)
    r.raise_for_status()
    from io import BytesIO

    df = pd.read_csv(BytesIO(r.content), sep=";", dtype=str, encoding="utf-8")
    df.columns = [c.strip().strip('"') for c in df.columns]
    return df


def main() -> int:
    log = setup_logging()
    ENTRADA.mkdir(parents=True, exist_ok=True)
    # limpa entradas anteriores
    for p in ENTRADA.glob("*.xlsx"):
        p.unlink()

    frames = []
    for ano, url in URLS.items():
        log.info("Baixando despesas %s...", ano)
        df = baixar(url)
        df["fonte_ano_arquivo"] = ano
        frames.append(df)
        log.info("  %d linhas, colunas=%s", len(df), list(df.columns)[:6])

    full = pd.concat(frames, ignore_index=True)
    # Injeta algumas linhas ruins de propósito para demonstrar rejeição
    bad = full.iloc[:3].copy()
    bad["Empenhado"] = "VALOR_INVALIDO"
    bad["Liquidado"] = "###"
    bad["Pago"] = "abc"
    bad["_injetado_para_demo"] = "sim"
    full = pd.concat([full, bad], ignore_index=True)

    # Divide por mês (simula arquivos mensais exportados por sistemas diferentes)
    if "Mês" not in full.columns:
        raise SystemExit("Coluna Mês não encontrada")
    full["Mês"] = full["Mês"].astype(str).str.strip().str.strip('"')
    full["Ano"] = full["Ano"].astype(str).str.strip().str.strip('"')

    grupos = full.groupby(["Ano", "Mês"], dropna=False)
    n_files = 0
    for (ano, mes), g in grupos:
        mes_z = str(mes).zfill(2)
        path = ENTRADA / f"despesas_{ano}_mes_{mes_z}.xlsx"
        g.drop(columns=["fonte_ano_arquivo"], errors="ignore").to_excel(path, index=False)
        n_files += 1
        log.info("Gerado %s (%d linhas)", path.name, len(g))

    meta = ROOT / "dados_entrada" / "FONTE.txt"
    meta.write_text(
        "Fonte: Portal de Dados Abertos da Prefeitura do Recife\n"
        "Conjunto: Despesas Orçamentárias — Despesas Órgão 2024 e 2025\n"
        "URLs:\n"
        + "\n".join(f"- {u}" for u in URLS.values())
        + "\nLicença: Open Data Commons ODbL\n"
        "Os arquivos .xlsx desta pasta são recortes mensais do CSV oficial,\n"
        "simulando exportações operacionais a consolidar.\n",
        encoding="utf-8",
    )
    log.info("Total de arquivos gerados: %d", n_files)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
