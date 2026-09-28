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
import subprocess
import sys
import unicodedata
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


def data_commit(path):
    """Data do último commit que mexeu no arquivo (0 se não der pra saber)."""
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%ct", "--", path.name],
            cwd=path.parent, capture_output=True, text=True, timeout=10,
        ).stdout.strip()
        return int(out) if out else 0
    except Exception:
        return 0


def abrir_planilha():
    """Abre a planilha mais recente de uploads/. Se houver mais de uma, tenta da
    mais nova para a mais antiga e pula arquivos que não abrem ou não têm a aba GERAL."""
    uploads_dir = REPO_ROOT / "uploads"
    candidatos = list(uploads_dir.glob("*.xlsm")) + list(uploads_dir.glob("*.xlsx"))
    if not candidatos:
        sys.exit("Nenhum arquivo .xlsm/.xlsx encontrado em uploads/")
    if len(candidatos) > 1:
        print("Aviso: mais de uma planilha em uploads/: " + ", ".join(p.name for p in candidatos))
    candidatos.sort(key=lambda p: (data_commit(p), p.stat().st_mtime), reverse=True)

    for caminho in candidatos:
        try:
            wb = openpyxl.load_workbook(caminho, data_only=True, keep_vba=True)
        except Exception as e:
            print(f"Ignorando {caminho.name}: não abriu como planilha ({e.__class__.__name__})")
            continue
        if "GERAL" not in wb.sheetnames:
            print(f"Ignorando {caminho.name}: não tem a aba GERAL")
            continue
        return caminho, wb
    sys.exit("Nenhuma planilha válida (com a aba GERAL) em uploads/")


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


ESTOQUE_MANUAL_PATH = DATA_DIR / "estoque.json"


def parse_estoque_planilha(wb):
    """Procura uma aba cujo nome contenha 'ESTOQUE' com as colunas
    MODELO (ou FILTRO) e QUANTIDADE (ou QUANT.). Retorna None se não achar."""
    aba = next((wb[n] for n in wb.sheetnames if "ESTOQUE" in n.upper()), None)
    if aba is None:
        return None

    # Acha a linha de cabeçalho e as colunas pelo nome (a ordem não importa).
    # Obrigatórias: MODELO e QUANT. | Opcionais: CÓDIGO, LOCAL, EST. MÍN, EST. MÁX
    cols = {}
    linha_cab = None
    for row in aba.iter_rows(min_row=1, max_row=10):
        achados = {}
        for c in row:
            campo = campo_do_cabecalho(c.value)
            if campo and campo not in achados:
                achados[campo] = c.column
        if "modelo" in achados and "quantidade" in achados:
            cols, linha_cab = achados, row[0].row
            break
    if not cols:
        print(f"Aviso: aba '{aba.title}' encontrada, mas sem colunas MODELO e QUANTIDADE")
        return None

    def valor(r, campo):
        return aba.cell(row=r, column=cols[campo]).value if campo in cols else None

    # O mesmo modelo pode aparecer em mais de uma linha: as quantidades são
    # SOMADAS; código, local, mínimo e máximo vêm da primeira linha preenchida.
    por_modelo = {}
    for r in range(linha_cab + 1, aba.max_row + 1):
        modelo = valor(r, "modelo")
        if modelo is None or str(modelo).strip() == "":
            continue
        nome = " ".join(str(modelo).split())
        chave = nome.upper().removesuffix(" SMC").strip()
        qtd = numero(valor(r, "quantidade"))
        extras = {
            "codigo": texto_limpo(valor(r, "codigo")),
            "local": texto_limpo(valor(r, "local")),
            "minimo": numero(valor(r, "minimo")),
            "maximo": numero(valor(r, "maximo")),
        }
        if chave not in por_modelo:
            por_modelo[chave] = {"modelo": nome, "quantidade": qtd, **extras}
            continue
        item = por_modelo[chave]
        if qtd is not None:
            item["quantidade"] = (item["quantidade"] or 0) + qtd
        for k, v in extras.items():
            if item.get(k) is None and v is not None:
                item[k] = v
    return list(por_modelo.values())


def campo_do_cabecalho(valor):
    """Traduz o texto do cabeçalho para o nome do campo (ignora acento e ponto)."""
    if valor is None:
        return None
    t = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode()
    t = " ".join(t.upper().replace(".", " ").split())
    if t in ("MODELO", "FILTRO"):
        return "modelo"
    if t.startswith("QUANT") or t == "QTD":
        return "quantidade"
    if t.startswith("COD"):
        return "codigo"
    if t.startswith("LOCAL"):
        return "local"
    if "MIN" in t.split() or t.startswith("EST MIN") or t in ("MINIMO", "ESTOQUE MINIMO"):
        return "minimo"
    if "MAX" in t.split() or t.startswith("EST MAX") or t in ("MAXIMO", "ESTOQUE MAXIMO"):
        return "maximo"
    return None


def texto_limpo(valor):
    """Código 31982 (número) vira '31982'; vazio vira None."""
    if valor is None or str(valor).strip() == "":
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def numero(valor):
    """Aceita 6, 6.0 ou o texto '6' digitado como texto no Excel."""
    if isinstance(valor, (int, float)):
        return int(valor) if float(valor).is_integer() else valor
    if isinstance(valor, str):
        try:
            v = float(valor.strip().replace(",", "."))
            return int(v) if v.is_integer() else v
        except ValueError:
            return None
    return None


def carregar_estoque(wb):
    """Planilha tem prioridade; se não tiver a aba de estoque, usa data/estoque.json."""
    estoque = parse_estoque_planilha(wb)
    if estoque is not None:
        print(f"Estoque lido da planilha: {len(estoque)} modelo(s)")
        return estoque, "planilha"
    if ESTOQUE_MANUAL_PATH.exists():
        estoque = json.loads(ESTOQUE_MANUAL_PATH.read_text(encoding="utf-8"))
        print(f"Estoque lido de data/estoque.json: {len(estoque)} modelo(s)")
        return estoque, "manual"
    return [], None


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
    excel_path, wb = abrir_planilha()
    print(f"Lendo: {excel_path.name}")

    itens = parse_geral(wb["GERAL"])
    estoque, origem_estoque = carregar_estoque(wb)

    today = datetime.now().strftime("%Y-%m-%d")

    DATA_DIR.mkdir(exist_ok=True)
    historico, eventos_novos = atualizar_historico(itens, today)

    payload = {
        "gerado_em": today,
        "arquivo_origem": excel_path.name,
        "itens": itens,
        "estoque_por_modelo": estoque,
        "estoque_origem": origem_estoque,
    }

    (DATA_DIR / "latest.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (DATA_DIR / f"{today}.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"OK: {len(itens)} itens e {len(estoque)} modelo(s) de estoque gravados em data/latest.json")
    print(f"Histórico: {eventos_novos} nova(s) troca(s) registrada(s), {len(historico)} evento(s) no total")


if __name__ == "__main__":
    main()
