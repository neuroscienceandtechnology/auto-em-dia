"""Baixa o levantamento semanal de preços da ANP e gera precos-combustiveis.json (preço médio por UF)."""
import io, json, re, sys, unicodedata, urllib.request
from datetime import datetime, timezone

PAGE = "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas"
UA = {"User-Agent": "AutoEmDia/1.0 (+https://github.com/neuroscienceandtechnology)"}

UFS = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM", "BAHIA": "BA", "CEARA": "CE",
    "DISTRITO FEDERAL": "DF", "ESPIRITO SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA", "MATO GROSSO": "MT",
    "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG", "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR",
    "PERNAMBUCO": "PE", "PIAUI": "PI", "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR", "SANTA CATARINA": "SC",
    "SAO PAULO": "SP", "SERGIPE": "SE", "TOCANTINS": "TO",
}
PRODUTOS = {"GASOLINA COMUM": "gasolina", "ETANOL HIDRATADO": "etanol", "OLEO DIESEL S10": "diesel", "GNV": "gnv"}


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
        return r.read()


def latest_xlsx():
    html = get(PAGE).decode("utf-8", "ignore")
    links = re.findall(r'href="([^"]*resumo_semanal[^"]*\.xlsx)"', html)
    if not links:
        sys.exit("Nenhuma planilha encontrada na página da ANP")

    def key(link):
        dates = re.findall(r"(\d{4})[-_](\d{2})[-_](\d{2})", link)
        return tuple(map(int, dates[-1])) if dates else (0, 0, 0)

    return max(links, key=key)


def parse(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if norm(n).startswith("ESTADO")), None)
    if sheet is None:
        sys.exit(f"Aba de estados não encontrada: {wb.sheetnames}")
    rows = list(sheet.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if any(norm(c) == "PRODUTO" for c in r))
    header = [norm(c) for c in rows[hi]]
    col = lambda name: next(i for i, h in enumerate(header) if h.startswith(name))
    c_uf, c_prod, c_preco = col("ESTADO"), col("PRODUTO"), col("PRECO MEDIO REVENDA")
    c_ini, c_fim = col("DATA INICIAL"), col("DATA FINAL")
    estados, inicio, fim = {}, None, None
    for r in rows[hi + 1:]:
        uf, prod = UFS.get(norm(r[c_uf])), PRODUTOS.get(norm(r[c_prod]))
        if not uf or not prod or not isinstance(r[c_preco], (int, float)):
            continue
        estados.setdefault(uf, {})[prod] = round(float(r[c_preco]), 2)
        inicio, fim = r[c_ini], r[c_fim]
    fmt = lambda d: d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)
    return {"fonte": "ANP — Levantamento de Preços de Combustíveis", "inicio": fmt(inicio), "fim": fmt(fim),
            "geradoEm": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), "estados": estados}


if __name__ == "__main__":
    url = latest_xlsx()
    if url.startswith("/"):
        url = "https://www.gov.br" + url
    print("Planilha:", url)
    result = parse(get(url))
    if len(result["estados"]) < 20:
        sys.exit(f"Poucos estados lidos ({len(result['estados'])}); formato da planilha mudou?")
    out = sys.argv[1] if len(sys.argv) > 1 else "precos-combustiveis.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1, sort_keys=True)
    print(json.dumps(result, ensure_ascii=False)[:600])
