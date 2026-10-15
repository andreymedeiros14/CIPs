# -*- coding: utf-8 -*-
"""Gera o dashboard de Health Score do FPSO Forte (HTML único, sem dependências externas).

Uso:
  python health_score/gerar_dashboard.py \
      --oleo Relatorio_Amostras_15_10_2026.xlsx \
      --vib "RDO SEMAPI - FPSO FORTE - VIB-AGO-26.xlsm" \
      --sap Base_notas_SAP.xlsx \
      --alarmes alarms-list.csv \
      --data-ref 2026-10-15 \
      --saida FPSO_Forte_Health_Score.html

O cálculo da nota é feito no navegador, a partir dos dados embutidos, para que os pesos
possam ser ajustados na própria página (painel "Metodologia").
"""
import argparse
import json
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config import ALARM_ASSETS, CASES, EQUIPMENT, SAP_SYSTEM_CODES, SYSTEMS  # noqa: E402

warnings.filterwarnings("ignore")

OIL_SHEETS = {"Óleo Lubrificante": "Óleo", "Diesel": "Diesel", "Líquido de Arrefecimento": "Arrefecimento"}
OIL_PARAMS = ["Fe", "Cu", "Pb", "Al", "Cr", "Si", "Na", "V40", "V100", "TAN", "ISO", "W", "PQI"]
OIL_STATUS = {"Normal": "N", "Monitorar": "M", "Crítico": "C"}


def norm(tag):
    return re.sub(r"[\s\-_]", "", str(tag or "")).upper()


def iso(d):
    if d is None or (isinstance(d, float) and pd.isna(d)) or pd.isna(d):
        return None
    if isinstance(d, str):
        return d[:10]
    return pd.Timestamp(d).strftime("%Y-%m-%d")


def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s in ("-", "nan", "None") else s


# --------------------------------------------------------------------------- óleo
def load_oil(path):
    rows = []
    for sheet, fluid in OIL_SHEETS.items():
        df = pd.read_excel(path, sheet_name=sheet, header=3, dtype=str)
        for _, r in df.iterrows():
            serie = clean(r.get("Série"))
            if not serie:
                continue
            params = {p: clean(r.get(p)) for p in OIL_PARAMS if clean(r.get(p))}
            rows.append(dict(
                s=serie, fr=clean(r.get("Frota")), c=clean(r.get("Nome do Compartimento")) or "(sem compartimento)",
                f=fluid, d=iso(r.get("Coleta")), st=OIL_STATUS.get(clean(r.get("Saúde")), "P"),
                n=clean(r.get("Número da Amostra")), os=clean(r.get("Ordem de serviço")),
                ol=clean(r.get("Tipo de Óleo")), p=params, tx=clean(r.get("Texto Interpretação")),
            ))
    return rows


# --------------------------------------------------------------------------- vibração
def load_vib(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["FPSO Forte"]
    header = None
    rows = []
    meta = {}
    for r in ws.iter_rows(values_only=True):
        r = list(r)
        if header is None:
            if r[1] == "Data de embarque:":
                meta["embarque"] = iso(r[2])
            elif r[1] == "Data de desembarque:":
                meta["desembarque"] = iso(r[2])
            elif r[1] == "Responsável pela coleta:":
                meta["resp"] = clean(r[2])
            elif r[2] == "Tag":
                header = r
            continue
        if clean(r[1]) == "Tipo":  # início da tabela-resumo no fim da aba
            break
        tag = clean(r[2])
        if not tag:
            continue
        desc = clean(r[3])
        rows.append(dict(
            tag=tag, k=norm(tag), area=clean(r[1]),
            desc=desc if norm(tag) in norm(desc) else "",
            grp=clean(r[4]), prev=clean(r[7]), cur=clean(r[8]), dt=iso(r[9]),
            just=clean(r[10]), obs=clean(r[11]),
        ))
    # colunas H e I = Jun/26 e Ago/26
    meta["col_prev"] = "Jun/26"
    meta["col_cur"] = "Ago/26"
    return rows, meta


# --------------------------------------------------------------------------- SAP
def load_sap(path):
    df = pd.read_excel(path)
    cols = list(df.columns)
    # A planilha tem três colunas "Descrição": local, texto breve da nota, equipamento.
    desc_cols = [c for c in cols if str(c).startswith("Descrição")]
    df = df.rename(columns={desc_cols[0]: "desc_loc", desc_cols[1]: "desc_nota", desc_cols[2]: "desc_eq"})
    closed = df["Status sistema"].astype(str).str.contains("MSEN") | df["Data encermto."].notna()
    total, n_closed = len(df), int(closed.sum())
    df = df[~closed]
    notes = []
    for _, r in df.iterrows():
        notes.append(dict(
            n=clean(r["Nota"]), t=clean(r["Tipo de nota"]), d=clean(r["desc_nota"]),
            tag=clean(r["Nº TAG"]), loc=clean(r["Loc.instalação"]), le=clean(r["desc_eq"]) or clean(r["desc_loc"]),
            cd=iso(r["Concl.desejada"]), dn=iso(r["Data da nota"]), pr=clean(r["Prioridade"]),
            st=clean(r["Status sistema"]), o=clean(r["Ordem"]), ct=clean(r["CenTrabPrinc."]),
        ))
    return notes, dict(total=total, encerradas=n_closed)


# --------------------------------------------------------------------------- alarmes
SEV_ALARM = {"Média": 1, "Alta": 2, "Muito Alta": 3}


def load_alarms(path):
    """Lê a lista de alarmes e mantém só os classificados como USEFUL."""
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig", index_col=False, dtype=str)
    total = len(df)
    df = df[df["Classification"].astype(str).str.strip().str.upper() == "USEFUL"]
    names, rows = {}, []
    for _, r in df.iterrows():
        asset = re.sub(r"^FORTE-", "", clean(r["Asset Name"]).split(" (")[0]).strip()
        name = re.sub(r"\s{2,}", " - ", clean(r["Name"]))
        names.setdefault(name, clean(r["Prognosis"]))
        rows.append(dict(
            a=asset, s=clean(r["Alarm Start Time"])[:16], e=clean(r["Alarm End Time"])[:16] or None,
            n=name, v=SEV_ALARM.get(clean(r["Severity"]), 1), g=int(clean(r["Number of Aggregated Alarms"]) or 0),
            m=clean(r["Method"]),
        ))
    starts = sorted(x["s"] for x in rows)
    meta = dict(total=total, useful=len(rows), first=starts[0][:10] if starts else None,
                last=starts[-1][:10] if starts else None)
    return rows, names, meta


def match_prefix(loc, prefix):
    return loc == prefix or loc.startswith(prefix + "-")


# --------------------------------------------------------------------------- montagem
def build(args):
    oil = load_oil(args.oleo)
    vib, vib_meta = load_vib(args.vib)
    notes, sap_meta = load_sap(args.sap)
    alarms, alarm_names, alarm_meta = load_alarms(args.alarmes)
    alarm_owner = {}
    for eid, assets in ALARM_ASSETS.items():
        for a in assets:
            alarm_owner.setdefault(a, []).append(eid)
    alarms_by_eq = {e["id"]: [] for e in EQUIPMENT}
    for al in alarms:
        for eid in alarm_owner.get(al["a"], []):
            alarms_by_eq[eid].append(al)

    eq_by_id = {e["id"]: e for e in EQUIPMENT}
    prefixes = sorted(((p, e["id"]) for e in EQUIPMENT for p in e["sap"]), key=lambda x: -len(x[0]))

    sap_by_eq = {e["id"]: [] for e in EQUIPMENT}
    sap_by_sys = {s["id"]: [] for s in SYSTEMS}
    for nt in notes:
        owner = next((eid for p, eid in prefixes if match_prefix(nt["loc"], p)), None)
        if owner:
            sap_by_eq[owner].append(nt)
        else:
            parts = nt["loc"].split("-")
            sysid = SAP_SYSTEM_CODES.get(parts[1]) if len(parts) > 1 else None
            if sysid:
                sap_by_sys[sysid].append(nt)

    vib_by_key = {}
    for v in vib:
        vib_by_key.setdefault(v["k"], []).append(v)
    oil_by_serie = {}
    for o in oil:
        oil_by_serie.setdefault(o["s"], []).append(o)

    used_vib, used_oil = set(), set()
    out_eq = []
    for e in EQUIPMENT:
        vrows = []
        for t in e["vib"]:
            k = norm(t)
            for v in vib_by_key.get(k, []):
                vrows.append({kk: vv for kk, vv in v.items() if kk != "k"})
                used_vib.add(k)
        # remove linhas duplicadas (mesma tag listada mais de uma vez na planilha)
        seen, vuniq = set(), []
        for v in vrows:
            key = (v["tag"], v["grp"], v["desc"], v["cur"], v["prev"])
            if key not in seen:
                seen.add(key)
                vuniq.append(v)
        orows = []
        for s in e["oil"]:
            orows.extend(oil_by_serie.get(s, []))
            if s in oil_by_serie:
                used_oil.add(s)
        orows.sort(key=lambda x: x["d"] or "")
        # mantém o texto de interpretação só nas 3 amostras mais recentes de cada compartimento
        cnt = {}
        slim = []
        for o in reversed(orows):
            key = (o["s"], o["c"])
            cnt[key] = cnt.get(key, 0) + 1
            o2 = dict(o)
            if cnt[key] > 3:
                o2.pop("tx", None)
            slim.append(o2)
        slim.reverse()
        sap = sorted(sap_by_eq[e["id"]], key=lambda x: x["cd"] or "9999")
        out_eq.append(dict(
            id=e["id"], sys=e["sys"], name=e["name"], tag=e["tag"], w=e["w"],
            sapPrefix=e["sap"], eng=e["eng"], cases=e["cases"], vib=vuniq, oil=slim, sap=sap,
            alm=sorted(alarms_by_eq[e["id"]], key=lambda x: x["s"]), almAssets=ALARM_ASSETS.get(e["id"], []),
        ))

    missing_cases = [c for e in EQUIPMENT for c in e["cases"] if c not in {x["id"] for x in CASES}]
    if missing_cases:
        raise SystemExit(f"Cases não cadastrados: {missing_cases}")

    unmapped_vib = sorted({v["tag"] for v in vib if v["k"] not in used_vib})
    unmapped_oil = sorted({o["s"] for o in oil if o["s"] not in used_oil})

    data = dict(
        refDate=args.data_ref,
        generatedAt=datetime.now().strftime("%Y-%m-%d %H:%M"),
        sources=dict(
            oleo=Path(args.oleo).name, vib=Path(args.vib).name, sap=Path(args.sap).name,
            alarmes=Path(args.alarmes).name,
            pptx="Apresentacao_FPSO_Forte.pptx", cases="Cases Forte / Cases Forte - pt2 (telas)",
        ),
        vibMeta=vib_meta, sapMeta=sap_meta, alarmMeta=alarm_meta, alarmNames=alarm_names,
        systems=SYSTEMS, equipment=out_eq, cases=CASES,
        sysNotes={k: sorted(v, key=lambda x: x["cd"] or "9999") for k, v in sap_by_sys.items()},
        coverage=dict(
            vibRows=len(vib), vibMapped=sum(1 for v in vib if v["k"] in used_vib), vibUnmapped=unmapped_vib,
            oilSamples=len(oil), oilMapped=sum(1 for o in oil if o["s"] in used_oil), oilUnmapped=unmapped_oil,
            sapOpen=len(notes), sapMapped=sum(len(v) for v in sap_by_eq.values()),
            sapSystem=sum(len(v) for v in sap_by_sys.values()),
            almUseful=len(alarms), almMapped=sum(1 for a in alarms if a["a"] in alarm_owner),
            almUnmapped=sorted({a["a"] for a in alarms if a["a"] not in alarm_owner}),
        ),
    )
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--oleo", required=True)
    ap.add_argument("--vib", required=True)
    ap.add_argument("--sap", required=True)
    ap.add_argument("--alarmes", required=True)
    ap.add_argument("--data-ref", default=datetime.now().strftime("%Y-%m-%d"))
    ap.add_argument("--saida", default="FPSO_Forte_Health_Score.html")
    ap.add_argument("--template", default=str(Path(__file__).resolve().parent / "template.html"))
    args = ap.parse_args()

    data = build(args)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = Path(args.template).read_text(encoding="utf-8").replace("/*__DATA__*/null", payload)
    Path(args.saida).write_text(html, encoding="utf-8")

    cov = data["coverage"]
    print(f"Equipamentos: {len(data['equipment'])} | Sistemas: {len(data['systems'])}")
    print(f"Vibração: {cov['vibMapped']}/{cov['vibRows']} linhas associadas")
    print(f"Óleo: {cov['oilMapped']}/{cov['oilSamples']} amostras associadas")
    print(f"SAP: {cov['sapMapped']} notas abertas em equipamentos, {cov['sapSystem']} só no sistema, "
          f"{cov['sapOpen']} abertas no total")
    print(f"Alarmes USEFUL: {cov['almMapped']}/{cov['almUseful']} associados. Sem equipamento: {cov['almUnmapped']}")
    print("Arquivo gerado:", args.saida)


if __name__ == "__main__":
    main()
