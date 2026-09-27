#!/usr/bin/env python3
"""
Lê o arquivo GESTÃO_TOLVA (.xlsm) e gera os JSONs usados pelo site.

Lógica replicada da planilha (aba GERAL, Tabela6):
  ULTIMA   = MAX(coluna "2025", coluna "OS")   -> data da última troca
  PRÓXIMA  = ULTIMA + FREQU.                   -> calculado no navegador (JS)
  DIAS RESTANTES = PRÓXIMA - HOJE              -> calculado no navegador (JS)
  STATUS   = <2 dias: PENDENTE | >=20 dias: DENTRO DO PRAZO | senão: ATENÇÃO

Como PRÓXIMA/DIAS RESTANTES/STATUS dependem da data de hoje, eles NÃO são
gravados no JSON — o site recalcula isso sozinho a cada carregamento.

Além disso, este script mantém um histórico de trocas (data/historico.json):
toda vez que a "ultima_troca" de um item muda em relação à execução anterior,
é registrado um evento de troca. Isso alimenta o gráfico de trocas por mês.
"""

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

EXCEL_EPOCH = datetime(1899, 12, 30)  # mesma referência de data serial do Excel

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"
HISTORICO_PATH = DATA_DIR / "historico.json"


def to_serial(value):
    if isinstance(value, datetime):
        return (value - EXCEL_EPOCH).days
    if isinstance(value, (int, float)):
        return value
    return None


def serial_to_iso(serial):
    return (EXCEL_EPOCH + timedelta(days=int(serial))).strftime("%Y-%m-%d")


def find_excel_file():
    uploads_dir = REPO_ROOT / "uploads"
    candidates = sorted(uploads_dir.glob("*.xlsm")) + sorted(uploads_dir.glob("*.xlsx"))
    if not candidates:
        sys.exit("Nenhum arquivo .xlsm/.xlsx encontrado em uploads/")
    return max(candidates, key=lambda p: p.stat().st_mtime)


def parse_geral(ws):
    """Lê a tabela Tabela6 (B2:K23) da aba GERAL."""
    itens = []
    row = 3
    while True:
        setor = ws.cell(row=row, column=2).value  # B
        if setor is None or str(setor).strip() == "":
            break

        tag = ws.cell(row=row, column=3).value        # C
        filtro = ws.cell(row=row, column=4).value      # D
        frequencia = ws.cell(row=row, column=5).value  # E
        col_2025 = ws.cell(row=row, column=6).value    # F
        col_os = ws.cell(row=row, column=7).value      # G

        serial_2025 = to_serial(col_2025)
        serial_os = to_serial(col_os)
        serials = [s for s in (serial_2025, serial_os) if s is not None]

        ultima_iso = serial_to_iso(max(serials)) if serials else None
        setor_nome = str(setor).strip()

        itens.append({
            "setor": setor_nome,
            "grupo": grupo_de(setor_nome),
            "tag": tag,
            "filtro": filtro,
            "frequencia_dias": frequencia,
            "ultima_troca": ultima_iso,
        })
        row += 1

    return itens


def grupo_de(setor_nome):
    """Extrai o nome do grupo/equipamento a partir do nome do setor.
    Ex: 'MESPACK A 01' -> 'MESPACK A' | 'LINHA PRINCIPAL 02' -> 'LINHA PRINCIPAL'
    Remove o número de posição no final (últimos 2 dígitos)."""
    partes = setor_nome.strip().split()
    if partes and partes[-1].isdigit():
        return " ".join(partes[:-1])
    return setor_nome


def parse_estoque(ws):
    estoque = []
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=3, max_col=5):
        quant, medida, descricao = (c.value for c in row)
        if quant is None and medida is None and descricao is None:
            continue
        estoque.append({
            "quantidade": quant,
            "medida": medida,
            "descricao": descricao,
        })
    return estoque


def atualizar_historico(itens_novos, hoje_iso):
    """Compara com o latest.json anterior (se existir) e registra trocas novas."""
    historico = []
    if HISTORICO_PATH.exists():
        try:
            historico = json.loads(HISTORICO_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            historico = []

    latest_path = DATA_DIR / "latest.json"
    itens_antigos_por_setor = {}
    if latest_path.exists():
        try:
            payload_antigo = json.loads(latest_path.read_text(encoding="utf-8"))
            itens_antigos_por_setor = {
                item["setor"]: item.get("ultima_troca") for item in payload_antigo.get("itens", [])
            }
        except json.JSONDecodeError:
            pass

    eventos_novos = 0
    for item in itens_novos:
        setor = item["setor"]
        ultima_nova = item["ultima_troca"]
        ultima_antiga = itens_antigos_por_setor.get(setor)
        # só registra troca se já existia um valor antes E o valor mudou
        # (evita registrar tudo na primeira execução do script)
        if ultima_antiga is not None and ultima_nova != ultima_antiga:
            historico.append({
                "data_troca": ultima_nova,
                "setor": setor,
                "grupo": item["grupo"],
                "filtro": item["filtro"],
                "registrado_em": hoje_iso,
            })
            eventos_novos += 1

    historico.sort(key=lambda e: e["data_troca"] or "")
    HISTORICO_PATH.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")
    return historico, eventos_novos


def main():
    excel_path = find_excel_file()
    print(f"Lendo: {excel_path.name}")

    wb = openpyxl.load_workbook(excel_path, data_only=True, keep_vba=True)

    itens = parse_geral(wb["GERAL"])
    estoque = parse_estoque(wb["Plan4"]) if "Plan4" in wb.sheetnames else []

    today = datetime.now().strftime("%Y-%m-%d")

    DATA_DIR.mkdir(exist_ok=True)
    historico, eventos_novos = atualizar_historico(itens, today)

    payload = {
        "gerado_em": today,
        "arquivo_origem": excel_path.name,
        "itens": itens,
        "estoque_filtros": estoque,
    }

    (DATA_DIR / "latest.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (DATA_DIR / f"{today}.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"OK: {len(itens)} itens e {len(estoque)} linhas de estoque gravados em data/latest.json")
    print(f"Histórico: {eventos_novos} nova(s) troca(s) registrada(s), {len(historico)} evento(s) no total")


if __name__ == "__main__":
    main()
