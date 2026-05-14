import pandas as pd
import json
import os
from datetime import datetime, date

EXCEL_PATH = r"C:\Users\ACER\Taalex Systemtechnik GmbH\Importação - Documentos\Importação Temu\Processo 2026\Controle Importação l Temu - 2026 - 2.xlsx"

OUTPUT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados_dashboard.json")

def safe_date(val):
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")
    return str(val)

def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()

def safe_num(val):
    try:
        if pd.isna(val):
            return 0
        return float(val)
    except:
        return 0

print("Lendo Excel...")
df = pd.read_excel(EXCEL_PATH, sheet_name="Follow Up", dtype=str)

df.columns = [c.strip() for c in df.columns]

df = df[df["Embarque"].notna() & (df["Embarque"].str.strip() != "") & (df["Embarque"].str.strip() != "0")]

embarques = []
hoje = datetime.today()

for _, row in df.iterrows():
    embarque = safe_str(row.get("Embarque"))
    projeto   = safe_str(row.get("Projeto"))
    cod       = safe_str(row.get("CodProjeto"))
    fornecedor= safe_str(row.get("Fornecedor"))
    destino   = safe_str(row.get("Destino"))
    status    = safe_str(row.get("Status"))
    prontidao = safe_str(row.get("Previsão de Prontidão"))
    qtd_cntr  = safe_str(row.get("Qty CNTR", ""))
    tipo_cntr = safe_str(row.get("Tipo CNTR", ""))
    obs       = safe_str(row.get("Observações", ""))

    etd_raw = row.get("ETD")
    eta_raw = row.get("ETA")

    try:
        etd_dt = pd.to_datetime(etd_raw, dayfirst=False, errors="coerce")
        etd_str = etd_dt.strftime("%d/%m/%Y") if pd.notna(etd_dt) else None
    except:
        etd_str = None

    try:
        eta_dt = pd.to_datetime(eta_raw, dayfirst=False, errors="coerce")
        eta_str = eta_dt.strftime("%d/%m/%Y") if pd.notna(eta_dt) else None
    except:
        eta_str = None

    etd_7dias = False
    if etd_str:
        try:
            etd_date = datetime.strptime(etd_str, "%d/%m/%Y")
            diff = (etd_date - hoje).days
            etd_7dias = 0 <= diff <= 7
        except:
            pass

    chegando_10 = False
    if eta_str:
        try:
            eta_date = datetime.strptime(eta_str, "%d/%m/%Y")
            diff = (eta_date - hoje).days
            chegando_10 = 0 <= diff <= 10
        except:
            pass

    if status.upper() == "FINALIZADO":
        status_op = "Finalizado"
    elif not prontidao and not etd_str:
        status_op = "Sem Prontidão e ETD"
    elif not prontidao:
        status_op = "Sem Prontidão"
    elif not etd_str:
        status_op = "Pronto - Aguardando ETD"
    else:
        status_op = "Prontidão e ETD OK"

    try:
        qtd = int(float(qtd_cntr)) if qtd_cntr else 0
    except:
        qtd = 0

    embarques.append({
        "embarque": embarque,
        "codProjeto": cod,
        "projeto": projeto,
        "fornecedor": fornecedor,
        "destino": destino,
        "qtdCntr": qtd,
        "tipoCntr": tipo_cntr,
        "prontidao": prontidao,
        "etd": etd_str,
        "eta": eta_str,
        "status": status,
        "statusOp": status_op,
        "etd7dias": etd_7dias,
        "chegando10": chegando_10,
        "observacoes": obs
    })

total_projetos = df["CodProjeto"].nunique()
ativos = [e for e in embarques if e["status"].upper() not in ["FINALIZADO"]]
total_containers = sum(e["qtdCntr"] for e in ativos if e["status"].upper() in ["SOBRE AGUAS", "SOBRE ÁGUAS"])
etd_7_count = sum(1 for e in embarques if e["etd7dias"])
chegando_10_count = sum(1 for e in embarques if e["chegando10"])

status_counts = {}
for e in embarques:
    s = e["status"]
    status_counts[s] = status_counts.get(s, 0) + 1

fornecedor_counts = {}
for e in ativos:
    for f in e["fornecedor"].split("+"):
        f = f.strip()
        if f:
            fornecedor_counts[f] = fornecedor_counts.get(f, 0) + 1

proximos_etd = sorted(
    [e for e in embarques if e["etd"] and e["status"].upper() != "FINALIZADO"],
    key=lambda x: datetime.strptime(x["etd"], "%d/%m/%Y")
)[:6]

proximos_eta = sorted(
    [e for e in embarques if e["eta"] and e["status"].upper() != "FINALIZADO"],
    key=lambda x: datetime.strptime(x["eta"], "%d/%m/%Y")
)[:6]

dados = {
    "geradoEm": hoje.strftime("%d/%m/%Y %H:%M"),
    "kpis": {
        "totalProjetos": total_projetos,
        "totalContainers": total_containers,
        "etd7dias": etd_7_count,
        "chegando10dias": chegando_10_count
    },
    "statusCounts": status_counts,
    "fornecedorCounts": fornecedor_counts,
    "proximosETD": proximos_etd,
    "proximosETA": proximos_eta,
    "embarques": embarques
}

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print(f"Dados gerados com sucesso!")
print(f"Total embarques: {len(embarques)}")
print(f"Total projetos: {total_projetos}")
print(f"ETD 7 dias: {etd_7_count}")
print(f"Arquivo salvo em: {OUTPUT_JSON}")
