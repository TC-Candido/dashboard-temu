# -*- coding: utf-8 -*-
import pandas as pd
import json
import os
import shutil
import tempfile
from datetime import datetime

EXCEL_PATH = r"C:\Users\ACER\Taalex Systemtechnik GmbH\Importação - Documentos\Importação Temu\Processo 2026\Controle Importação l Temu - 2026 - 2.xlsx"
OUTPUT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados_cambio.json")
OUTPUT_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_cambio.html")
TEMPLATE_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_cambio_template.html")

def safe_str(val):
    try:
        if pd.isna(val) or val is None:
            return ""
        return str(val).strip()
    except:
        return ""

def safe_num(val):
    try:
        if pd.isna(val):
            return 0
        return float(val)
    except:
        return 0

def dias_para(date_val, hoje):
    if date_val is None:
        return None
    try:
        if pd.isna(date_val):
            return None
    except:
        pass
    try:
        if isinstance(date_val, str):
            d = datetime.strptime(date_val, "%d/%m/%Y")
        else:
            d = pd.to_datetime(date_val).to_pydatetime()
        if d.year < 2000:
            return None
        diff = (d.replace(hour=0,minute=0,second=0,microsecond=0) - hoje).days
        return diff
    except:
        return None

def formatar_dias(diff):
    if diff is None:
        return None
    if diff < 0:
        return f"{abs(diff)}d atrás"
    if diff == 0:
        return "Hoje"
    if diff == 1:
        return "Amanhã"
    return f"{diff} dias"

def status_venc(saldo, previsao_diff):
    if saldo <= 0:
        return "Quitado"
    if previsao_diff is None:
        return "Sem previsão"
    if previsao_diff < 0:
        return "Vencido"
    if previsao_diff <= 30:
        return "Vence em 30 dias"
    return "Em dia"

print("Lendo Excel...")

tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
tmp.close()
try:
    shutil.copy2(EXCEL_PATH, tmp.name)
    df = pd.read_excel(tmp.name, sheet_name="Cambio Fornecedor", header=0, dtype=str)
finally:
    try:
        os.unlink(tmp.name)
    except:
        pass

df.columns = [str(c).strip() if str(c) != 'nan' else f"Col_{i}" for i, c in enumerate(df.columns)]

hoje = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

cambios = []
for _, row in df.iterrows():
    projeto    = safe_str(row.get("Projeto"))
    cod        = safe_str(row.get("CodProjeto"))
    embarque   = safe_str(row.get("Embarque"))
    fornecedor = safe_str(row.get("Fornecedor"))
    fatura     = safe_str(row.get("Fatura"))

    if not projeto or not fornecedor:
        continue

    def get_num(col):
        return safe_num(row.get(col, 0))

    valor  = get_num("Valor")
    pct1   = get_num("1ª Parcela - %")
    pct2   = get_num("2ª Parcela - %")
    pct3   = get_num("3ª Parcela - %")
    saldo  = get_num("Saldo a Pagar Fabrica")
    status_c = safe_str(row.get("Status Câmbio", ""))

    # Pago parcelas — colunas sem nome
    cols = list(df.columns)
    def get_col_after(name):
        try:
            idx = cols.index(name)
            return safe_num(row.iloc[idx + 1])
        except:
            return 0

    pago1 = get_col_after("1ª Parcela - %")
    pago2 = get_col_after("2ª Parcela - %")
    pago3 = get_col_after("3ª Parcela - %")

    prev_raw = row.get("Previsão.Pagamento Final", None)
    prev_diff = dias_para(prev_raw, hoje)
    prev_str = None
    if prev_raw and str(prev_raw) not in ['nan', 'None', '']:
        try:
            prev_str = pd.to_datetime(prev_raw).strftime("%d/%m/%Y")
            if pd.to_datetime(prev_raw).year < 2000:
                prev_str = None
        except:
            pass

    sv = status_venc(saldo, prev_diff)

    cambios.append({
        "projeto":    projeto,
        "codProjeto": cod,
        "embarque":   "" if embarque in ["0"] else embarque,
        "fornecedor": fornecedor,
        "fatura":     fatura,
        "valor":      round(valor, 2),
        "pct1":       round(pct1 * 100, 0) if pct1 <= 1 else round(pct1, 0),
        "pago1":      round(pago1, 2),
        "pct2":       round(pct2 * 100, 0) if pct2 <= 1 else round(pct2, 0),
        "pago2":      round(pago2, 2),
        "pct3":       round(pct3 * 100, 0) if pct3 <= 1 else round(pct3, 0),
        "pago3":      round(pago3, 2),
        "saldo":      round(saldo, 2),
        "statusCambio": status_c,
        "previsao":   prev_str,
        "previsaoDias": prev_diff,
        "previsaoLabel": formatar_dias(prev_diff),
        "statusVenc": sv
    })

total_saldo = sum(c["saldo"] for c in cambios if c["saldo"] > 0)
qtd_aberto  = sum(1 for c in cambios if c["saldo"] > 0)
venc_30     = sum(c["saldo"] for c in cambios if c["saldo"] > 0 and c["statusVenc"] == "Vence em 30 dias")
vencidos    = sum(c["saldo"] for c in cambios if c["saldo"] > 0 and c["statusVenc"] == "Vencido")

forn_saldo = {}
for c in cambios:
    if c["saldo"] > 0:
        for f in c["fornecedor"].split("+"):
            f = f.strip()
            if f:
                forn_saldo[f] = round(forn_saldo.get(f, 0) + c["saldo"], 2)
forn_saldo = dict(sorted(forn_saldo.items(), key=lambda x: x[1], reverse=True))

status_counts = {}
for c in cambios:
    if c["saldo"] > 0:
        s = c["statusCambio"] or "Sem status"
        status_counts[s] = status_counts.get(s, 0) + 1

proximos = sorted(
    [c for c in cambios if c["saldo"] > 0 and c["previsaoDias"] is not None],
    key=lambda x: x["previsaoDias"]
)[:8]

dados = {
    "geradoEm":    datetime.now().strftime("%d/%m/%Y %H:%M"),
    "kpis": {
        "totalSaldo": round(total_saldo, 2),
        "qtdAberto":  qtd_aberto,
        "venc30":     round(venc_30, 2),
        "vencidos":   round(vencidos, 2)
    },
    "fornSaldo":    forn_saldo,
    "statusCounts": status_counts,
    "proximos":     proximos,
    "cambios":      cambios
}

# Salva JSON
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

# Embutir dados no HTML
if os.path.exists(TEMPLATE_HTML):
    with open(TEMPLATE_HTML, "r", encoding="utf-8") as f:
        html = f.read()
    dados_js = json.dumps(dados, ensure_ascii=False)
    html = html.replace("var DADOS = DADOS_JSON;", f"var DADOS = {dados_js};")
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML atualizado com dados embutidos!")

print(f"Dados gerados com sucesso!")
print(f"Total câmbios: {len(cambios)}")
print(f"Em aberto: {qtd_aberto}")
print(f"Total saldo: $ {total_saldo:,.2f}")
print(f"Vencidos: $ {vencidos:,.2f}")
print(f"JSON salvo: {OUTPUT_JSON}")
