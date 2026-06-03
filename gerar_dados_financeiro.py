# -*- coding: utf-8 -*-
import pandas as pd
import json
import os
import shutil
import tempfile
from datetime import datetime

EXCEL_PATH = r"C:\Users\ACER\Taalex Systemtechnik GmbH\Importação - Documentos\Importação Temu\Processo 2026\Controle Importação l Temu - 2026 - 2.xlsx"
OUTPUT_JSON  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados_financeiro.json")
OUTPUT_HTML  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_financeiro.html")
TEMPLATE_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_financeiro_template.html")

def safe_str(val):
    try:
        if pd.isna(val): return ""
    except: pass
    return str(val).strip()

def safe_num(val):
    try:
        if pd.isna(val): return 0
        return float(str(val).replace(",", "."))
    except: return 0

def parse_date(val):
    if val is None: return None
    try:
        if pd.isna(val): return None
    except: pass
    try:
        d = pd.to_datetime(val, dayfirst=False)
        if d.year < 2000: return None
        return d
    except: return None

print("Lendo Excel - aba Financeiro...")
tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
tmp.close()
try:
    shutil.copy2(EXCEL_PATH, tmp.name)
    df = pd.read_excel(tmp.name, sheet_name="Financeiro", header=0, dtype=str)
finally:
    try: os.unlink(tmp.name)
    except: pass

df.columns = [str(c).strip() if str(c) != "nan" else f"Col_{i}" for i, c in enumerate(df.columns)]
df = df[df["Projeto"].notna() & (df["Projeto"].str.strip() != "") & (df["Projeto"].str.strip() != "nan")]

hoje = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

projetos = []
for _, row in df.iterrows():
    projeto   = safe_str(row.get("Projeto", ""))
    cod       = safe_str(row.get("CodProjeto", ""))
    embarque  = safe_str(row.get("Embarque", ""))
    status    = safe_str(row.get("STATUS", ""))
    cif       = safe_num(row.get("CIF", 0))
    impostos  = safe_num(row.get("IMPOSTOS", 0))
    despesas  = safe_num(row.get("DESPESAS", 0))
    frete     = safe_num(row.get("Frete", 0))
    tot_cambio  = safe_num(row.get("TOTAL DESEMBOLSO COM CÂMBIO", 0))
    tot_chegada = safe_num(row.get("Total de pagamento chegada", 0))

    prev_dt = parse_date(row.get("Previsão de Chegada", None))
    prev_str = prev_dt.strftime("%d/%m/%Y") if prev_dt else None
    dias_chegada = None
    semana_label = None
    if prev_dt:
        diff = (prev_dt - hoje).days
        dias_chegada = diff
        if 0 <= diff <= 7:
            semana_label = "Esta semana"
        elif 0 <= diff <= 15:
            semana_label = "Próximos 15 dias"
        elif 0 <= diff <= 30:
            semana_label = "Este mês"
        elif diff < 0:
            semana_label = "Atrasado"
        else:
            semana_label = "Futuro"

    if not projeto or (cif == 0 and tot_chegada == 0):
        continue

    projetos.append({
        "projeto":      projeto,
        "codProjeto":   cod,
        "embarque":     embarque,
        "status":       status,
        "cif":          round(cif, 2),
        "impostos":     round(impostos, 2),
        "despesas":     round(despesas, 2),
        "frete":        round(frete, 2),
        "totalCambio":  round(tot_cambio, 2),
        "totalChegada": round(tot_chegada, 2),
        "previsao":     prev_str,
        "diasChegada":  dias_chegada,
        "semanaLabel":  semana_label,
    })

# KPIs
total_geral      = sum(p["totalChegada"] for p in projetos)
total_semana     = sum(p["totalChegada"] for p in projetos if p["semanaLabel"] == "Esta semana")
total_quinzena   = sum(p["totalChegada"] for p in projetos if p["semanaLabel"] in ["Esta semana", "Próximos 15 dias"])
total_mes        = sum(p["totalChegada"] for p in projetos if p["semanaLabel"] in ["Esta semana", "Próximos 15 dias", "Este mês"])
qtd_chegando     = sum(1 for p in projetos if p["diasChegada"] is not None and p["diasChegada"] >= 0)

# Composição média (para gráfico)
comp = {
    "impostos": round(sum(p["impostos"] for p in projetos), 2),
    "despesas": round(sum(p["despesas"] for p in projetos), 2),
    "frete":    round(sum(p["frete"] for p in projetos), 2),
}

dados = {
    "geradoEm": datetime.now().strftime("%d/%m/%Y %H:%M"),
    "kpis": {
        "totalGeral":    round(total_geral, 2),
        "totalSemana":   round(total_semana, 2),
        "totalQuinzena": round(total_quinzena, 2),
        "totalMes":      round(total_mes, 2),
        "qtdChegando":   qtd_chegando,
    },
    "composicao": comp,
    "projetos": sorted(projetos, key=lambda x: (x["diasChegada"] if x["diasChegada"] is not None else 9999))
}

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

if os.path.exists(TEMPLATE_HTML):
    with open(TEMPLATE_HTML, "r", encoding="utf-8") as f:
        html = f.read()
    dados_js = json.dumps(dados, ensure_ascii=False)
    html = html.replace("var DADOS = __DADOS_JSON__;", f"var DADOS = {dados_js};")
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML atualizado!")

print(f"Projetos: {len(projetos)}")
print(f"Total a desembolsar: R$ {total_geral:,.2f}")
print(f"Esta semana: R$ {total_semana:,.2f}")
print(f"Quinzena: R$ {total_quinzena:,.2f}")
print(f"Mês: R$ {total_mes:,.2f}")
