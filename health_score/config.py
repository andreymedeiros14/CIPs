# -*- coding: utf-8 -*-
"""Cadastro de sistemas, equipamentos, cases e avaliação do engenheiro (FPSO Forte).

Cada equipamento aponta para as chaves usadas em cada fonte de dados:
  sap   -> prefixos de Local de Instalação (Base notas SAP). Vale o prefixo mais longo.
  vib   -> TAGs da aba "FPSO Forte" do RDO de vibração (comparação ignora espaços e hífens).
  oil   -> coluna "Série" do Relatório de Amostras (óleo, diesel e arrefecimento).
  cases -> IDs dos cases (telas "Cases Forte").
  eng   -> avaliação do engenheiro de manutenção (Apresentacao_FPSO_Forte.pptx).
           st: D = disponível, R = com restrição, I = indisponível.
           pend: True quando o equipamento está disponível mas com pendência registrada.
  w     -> peso de criticidade do equipamento dentro do sistema (1 a 5).
"""

# Peso de cada sistema no Health Score global e justificativa.
# Pesos fixos definidos pelo usuário (calibração de 15/10/2026). Peso 0 tira o sistema do global.
SYSTEMS = [
    dict(id="GP", name="Geração Principal de Energia", w=8, icon="power",
         why="Turbogeradores alimentam toda a planta. Perda de geração leva a shutdown parcial ou total."),
    dict(id="CG", name="Compressão de Gás", w=8, icon="compressor",
         why="Sem compressão o gás associado vai para o flare e a produção de óleo é restringida pelo limite de queima."),
    dict(id="CI", name="Combate a Incêndio", w=7, icon="fire",
         why="Elemento crítico de segurança (SCE). Indisponibilidade pode exigir parada da unidade."),
    dict(id="GE", name="Geração de Emergência", w=1, icon="power",
         why="Elemento crítico de segurança. Alimenta cargas essenciais em blackout."),
    dict(id="IA", name="Injeção de Água", w=1, icon="pump",
         why="Manutenção da pressão do reservatório. Impacto direto na curva de produção."),
    dict(id="CA", name="Captação de Água do Mar", w=1, icon="wave",
         why="Fornece água para resfriamento e alimenta a injeção. Falha afeta vários sistemas."),
    dict(id="RF", name="Água de Resfriamento", w=1, icon="cool",
         why="Resfriamento de turbogeradores e compressores. Falha provoca trip por temperatura."),
    dict(id="EL", name="Sistema Elétrico, Transformadores e UPS", w=1, icon="bolt",
         why="Transformadores dos compressores e UPS de sistemas de segurança e controle."),
    dict(id="EP", name="Escoamento de Petróleo", w=0, icon="oil",
         why="Transferência do óleo tratado para os tanques de carga."),
    dict(id="TO", name="Tratamento de Óleo", w=0.5, icon="process",
         why="Separação e aquecimento do óleo. Afeta especificação e vazão de produção."),
    dict(id="CP", name="Controle de Poços (HPU Submarina)", w=0.5, icon="valve",
         why="Comando das válvulas submarinas. Perda de HPU fecha poços."),
    dict(id="AR", name="Ar Comprimido", w=0.5, icon="air",
         why="Ar de instrumento para válvulas de controle e ESD."),
    dict(id="AQ", name="Água Quente (Cooling Medium)", w=0, icon="heat",
         why="Meio de aquecimento do processo. Possui redundância."),
    dict(id="AP", name="Tratamento de Água Produzida", w=0.5, icon="water",
         why="Enquadramento e descarte ou reinjeção da água produzida. Restrição pode limitar produção."),
    dict(id="TC", name="Telecomunicações", w=0.5, icon="radio",
         why="PAGA, rádio e GMDSS são itens de segurança. Demais itens são de suporte."),
    dict(id="GD", name="Movimentação de Cargas (Guindastes)", w=0.5, icon="crane",
         why="Logística de bordo e operações de carga. Não para a produção de imediato."),
    dict(id="MA", name="Unidades Hidráulicas, Offloading e Marinha", w=0.5, icon="ship",
         why="HPUs de carga, lastro, offloading e amarração."),
    dict(id="SV", name="Salvatagem (Baleeiras e Bote de Resgate)", w=4, icon="boat",
         why="Elemento crítico de segurança para abandono da unidade e resgate de homem ao mar."),
    dict(id="OD", name="Óleo Diesel", w=0, icon="drop",
         why="Abastecimento de diesel para geração de emergência, auxiliar e partida. Possui redundância."),
    dict(id="GA", name="Geração Auxiliar", w=0.5, icon="power",
         why="Apoio à geração principal. Baixo impacto direto na produção."),
    dict(id="HV", name="HVAC e Água Gelada", w=0.5, icon="fan",
         why="Climatização de salas técnicas e acomodações. Possui redundância."),
]

# Mapeamento do 2º nível do Local de Instalação SAP para o sistema (notas sem equipamento).
SAP_SYSTEM_CODES = {
    "FSGP": "GP", "FSGE": "GE", "FSGA": "GA", "PSCG": "CG", "PURV": "CG", "PSAI": "IA",
    "FSCA": "CA", "FSCI": "CI", "FECC": "EL", "FEIE": "EL", "FETL": "EL", "FSTI": "TC",
    "FSMC": "GD", "FSRA": "RF", "FSAQ": "AQ", "PSEP": "EP", "PSTP": "TO", "PSTA": "AP",
    "PSCP": "CP", "FSAC": "AR", "FSAP": "AR", "FSAR": "HV", "PSOF": "MA", "FSSV": "SV", "FSOD": "OD",
}


def E(id, sys, name, tag, w, sap=(), vib=(), oil=(), cases=(), eng=None):
    return dict(id=id, sys=sys, name=name, tag=tag, w=w, sap=list(sap), vib=list(vib),
                oil=list(oil), cases=list(cases), eng=eng)


def eng(st, obs="", os="", loc="", pend=False):
    return dict(st=st, obs=obs, os=os, loc=loc, pend=pend)


EQUIPMENT = [
    # ---------------- Geração Principal ----------------
    E("TGA", "GP", "Turbogerador A", "GE-514001A", 5, sap=["FORTE-FSGP-TGA"],
      vib=["GE-514001A", "TS-GE-514001A", "GB-GE-514001A", "M-B-GE-514001A-1A", "M-B-GE-514001A-1B"],
      oil=["TQ-GE-514001A-01", "TQ-GE-514001A-02", "TQ-GE-514001A-03", "TP-GE-514001A-01",
           "TP-GE-514001A-03", "TQ-GE-524001A-03", "TQ-GE-514001-04-A"],
      cases=["20260027"],
      eng=eng("D", "L1 em atraso. Baker a bordo 14/10", "4326224 (OSs com prontidão e preventivas por oportunidade)",
              "FORTE-FSGP-TGA-01-10-01", pend=True)),
    E("TGB", "GP", "Turbogerador B", "GE-514001B", 5, sap=["FORTE-FSGP-TGB"],
      vib=["M-B-GE-514001B-1A", "M-B-GE-514001B-1B"],
      oil=["TQ-GE-514001B-01", "TQ-GE-514001B-02", "TQ-GE-514001B-03"],
      eng=eng("I", "TG-B em comissionamento (P-90: 02/27)", "Diversas", "FORTE-FSGP-TGB-01-10-01")),
    E("TGC", "GP", "Turbogerador C", "GE-514001C", 5, sap=["FORTE-FSGP-TGC"],
      vib=["GE-514001C", "TS-GE-514001C", "GB-GE-514001C", "M-B-GE-514001C-1A", "M-B-GE-514001C-1B"],
      oil=["TQ-GE-514001C-01", "TQ-GE-514001C-02", "TQ-GE-514001C-03", "TP-GE-514001C-01",
           "TP-GE-514001C-03", "TQ-GE-524001C-02", "CIP-F2081", "GTGC_FORTE", "TQ-GE-514001-04-C"],
      cases=["20260061", "20260024", "20250039"],
      eng=eng("D", "L1 em atraso. Embarque da Baker 15/10. L3 em atraso (previsão 09/27). Óleo degradado",
              "4330029 (OSs com prontidão e preventivas por oportunidade) / 4288314 (óleo)",
              "FORTE-FSGP-TGC-01-10-01", pend=True)),
    E("TGD", "GP", "Turbogerador D", "GE-514001D", 5, sap=["FORTE-FSGP-TGD"],
      vib=["GE-514001D", "TS-GE-514001D", "GB-GE-514001D", "M-B-GE-514001D-1A", "M-B-GE-514001D-1B"],
      oil=["TQ-GE-514001D-01", "TQ-GE-514001D-02", "TQ-GE-514001D-03", "TP-GE-514001D-01",
           "TP-GE-514001D-03", "TQ-GE-524001D-23", "TQ-GE-524001", "TQ-GE-514001-04-D", "TANQUECOMBUSTTGD"],
      cases=["20260064", "20250019"],
      eng=eng("R", "TG-D sem válvula de líquido comissionada. Recomendações da L1 (15/09) pendentes",
              "Criar OS", "FORTE-FSGP-TGD-01-10-01")),

    E("SKAB", "GP", "Skid de Partida Hidráulica TG A/B", "TQ-GE-514001AB-01A", 2, sap=["FORTE-FSGP-HPU-01"],
      oil=["TG-GE-5140001AB-01A", "TG-GE-5140001AB-01B", "TG-GE-514001AB-01A", "TP-GE-514001AB-01A",
           "TP-GE-514001AB-01B", "TQ-GE-514001AB-01A"]),
    E("SKCD", "GP", "Skid de Partida Hidráulica TG C/D", "TQ-GE-514001CD-01A", 2, sap=["FORTE-FSGP-HPU-04"],
      oil=["TG-GE-5140001CD-01A", "TG-GE-5140001CD-01B", "TP-GE-514001CD-01A", "TP-GE-514001CD-01B"]),

    # ---------------- Geração de Emergência / Auxiliar ----------------
    E("GEE", "GE", "Gerador de Emergência", "GE-514003", 5, sap=["FORTE-FSGE"],
      vib=["GE-514003", "MC-GE-514003"], oil=["GE-514003", "GE-514003A"],
      eng=eng("D", "", "", "FORTE-FSGE-MGE-01-01")),
    E("GEA", "GA", "Gerador Auxiliar", "GE-514002", 4, sap=["FORTE-FSGA"],
      vib=["GE-514002", "MC-GE-514002"], oil=["MC-GE-514002", "TQ-MC-GE-514002-02"],
      eng=eng("D", "", "", "FORTE-FSGA-MGE-01-01")),

    # ---------------- Compressão de Gás ----------------
    E("MCA", "CG", "Compressor Principal A (MC-A)", "UC-123101A", 5, sap=["FORTE-PSCG-GAS-01"],
      vib=["M-C-UC-123101A", "W-C-UC-123101A", "C-UC-123101A-01", "C-UC-123101A-02"],
      oil=["TQ-Z-UC-1231-01A", "UC-123101A"],
      eng=eng("I", "Troca dos bundles LP/IP e HP, aguardando reparo dos bundles (11/26). "
                   "Ensaios elétricos e instalação de para-raio pendentes",
              "4391314 / 4390456 (RR) / 4390455 (RR)", "FORTE-PSCG-GAS-01")),
    E("MCB", "CG", "Compressor Principal B (MC-B)", "UC-123101B", 5, sap=["FORTE-PSCG-GAS-04"],
      vib=["M-C-UC-123101B", "W-C-UC-123101B", "C-UC-123101B-01", "C-UC-123101B-02"],
      oil=["TQ-Z-UC-1231-01B", "UC-123101B"], cases=["20260059", "20260049"],
      eng=eng("D", "", "", "FORTE-PSCG-GAS-04")),
    E("MCC", "CG", "Compressor Principal C (MC-C)", "UC-123101C", 5, sap=["FORTE-PSCG-GAS-07"],
      vib=["M-C-UC-123101C", "W-C-UC-123101C", "C-UC-123101C-01", "C-UC-123101C-02"],
      oil=["TQ-Z-UC-1231-01C", "UC-123101C", "TQ-C-UE-123101C"], cases=["20260038", "20250040"],
      eng=eng("R", "Vibração elevada no conjunto HP. Troca dos bundles em 03/2027", "Criar OS",
              "FORTE-PSCG-GAS-07")),
    E("BOO", "CG", "Compressor Booster (UC-1225)", "C-UC-122501", 4, sap=["FORTE-PURV"],
      vib=["C-UC-122501-01", "M-C-UC-122501"], oil=["C-UC-122501", "C-UC-1225001"], cases=["20260026"],
      eng=eng("I", "Trocador de MLO em reparo (RR)", "Criar OS", "FORTE-PURV-URV-01-01-01")),
] + [
    E(f"DGS{t}{n}", "CG", f"Compressor N2 DGS {t}-0{n}", f"C-GG-C-123101{t}-01{'A' if n == 1 else 'B'}", 2,
      sap=[f"FORTE-PSCG-GAS-0{ {'A': 1, 'B': 4, 'C': 7}[t] }-43-{'01' if n == 1 else '10'}"],
      vib=[f"C-GG-C-123101{t}-01{'A' if n == 1 else 'B'}", f"C-GG-UC-123101{t}-01{'A' if n == 1 else 'B'}"],
      oil=[f"C-GG-C-123101{t}-01{'A' if n == 1 else 'B'}"],
      cases=["20250025"] if t == "C" else [],
      eng=eng("D", "", "", f"FORTE-PSCG-GAS-0{ {'A': 1, 'B': 4, 'C': 7}[t] }-43-{'01' if n == 1 else '10'}-04"))
    for t in "ABC" for n in (1, 2)
] + [
    # ---------------- Injeção de Água ----------------
    E("BIA", "IA", "Bomba de Injeção A", "B-UB-125101A", 4, sap=["FORTE-PSAI-MBB-01-01"],
      vib=["B-UB-125101A", "M-B-UB-125101A"]),
    E("BIB", "IA", "Bomba de Injeção B", "B-UB-125101B", 4, sap=["FORTE-PSAI-MBB-01-07"],
      vib=["B-UB-125101B", "M-B-UB-125101B"], oil=["B-UB-125101B", "P-M-UB-125101B", "TQ-B-UB-125101B"],
      cases=["20260052"], eng=eng("R", "Alta vibração", "", "FORTE-PSAI-MBB-01-07")),
    E("BIC", "IA", "Bomba de Injeção C", "B-UB-125101C", 4, sap=["FORTE-PSAI-MBB-01-10"],
      vib=["B-UB-125101C", "M-B-UB-125101C"], oil=["B-UB-125101C-01", "P-M-UB-125101C", "TQ-B-UB-125101C"],
      cases=["20260050", "20260035"], eng=eng("D", "", "", "FORTE-PSAI-MBB-01-10")),
    E("BID", "IA", "Bomba de Injeção D", "B-UB-125101D", 4, sap=["FORTE-PSAI-MBB-01-13"],
      vib=["B-UB-125101D", "M-B-UB-125101D"], oil=["B-UB-125101D", "P-M-UB-125101D", "TQ-B-UB-125101D"],
      cases=["20260047"], eng=eng("D", "", "", "FORTE-PSAI-MBB-01-13")),
    E("BTA", "IA", "Bombas Booster de Injeção, Trem A (01A/01B/01C)", "B-UT-125101A", 2,
      sap=["FORTE-PSAI-URS-01-13"],
      vib=[f"{p}B-UT-125101A-01{x}" for x in "ABC" for p in ("", "M-")]),
    E("BTB", "IA", "Bombas Booster de Injeção, Trem B (01A/01B/01C)", "B-UT-125101B", 2,
      sap=["FORTE-PSAI-URS-01-16"],
      vib=[f"{p}B-UT-125101B-01{x}" for x in "ABC" for p in ("", "M-")]),
    E("LPA", "IA", "Bomba de Água do Mar p/ Injeção A", "B-511104A", 3,
      sap=["FORTE-PSAI-MBB-04-01", "FORTE-PSAI-MBB-04-04"], oil=["B-511104A", "UH-B-511104A"]),
    E("LPB", "IA", "Bomba de Água do Mar p/ Injeção B", "B-511104B", 3,
      sap=["FORTE-PSAI-MBB-04-07", "FORTE-PSAI-MBB-04-10"], oil=["B-511104B", "UH-B-511104B"]),
] + [
    # ---------------- Captação de Água do Mar ----------------
    E(f"CA{x}", "CA", f"Bomba de Captação {x}", f"B-511101{x}", 4, sap=[f"FORTE-FSCA-MBB-{l}"],
      vib=[f"B-511101{x}", f"M-B-511101{x}"], eng=eng("D", "", "", f"FORTE-FSCA-MBB-{l}"))
    for x, l in zip("ABCD", ("01", "04", "07", "10"))
] + [
    E("CAE", "CA", "Bomba de Captação de Emergência", "B-511102", 3, sap=["FORTE-FSCA-MBB-13"],
      vib=["B-511102", "M-B-511102"]),

    # ---------------- Combate a Incêndio ----------------
    E("FPA", "CI", "Bomba de Incêndio A", "B-UB-542001A", 5, sap=["FORTE-FSCI-RAC-01"],
      oil=["MC-UB-542001A", "TQ-UB-542001A-02", "MCB-UB-542001-02A"],
      eng=eng("D", "", "", "FORTE-FSCI-RAC-01")),
    E("FPB", "CI", "Bomba de Incêndio B", "B-UB-542001B", 5, sap=["FORTE-FSCI-RAC-07"],
      oil=["MC-UB-542001B", "B-UB-542001B-02", "BUB-542001B-02", "TQ-UB-542001B-02", "MCB-UB-542001-02B", "8EM00495"],
      eng=eng("D", "", "", "FORTE-FSCI-RAC-07")),
    E("FPC", "CI", "Bomba de Incêndio C", "B-UB-542001C", 5, sap=["FORTE-FSCI-RAC-13"],
      oil=["MC-UB-542001C", "TQ-UB-542001C-01", "TQ-UB-542001C-02", "MCB-UB-542001-02C"],
      eng=eng("R", "Possível vazamento no selo da bomba. Overhaul programado para 2027", "Criar OS",
              "FORTE-FSCI-RAC-13")),
] + [
    E(f"HU{x}", "CI", f"Unidade Hidráulica da Bomba de Incêndio {x}", f"UB-542001{x}", 3,
      sap=[f"FORTE-FSCI-RAC-{l}"], oil=[f"TP-UB-542001{x}-02"], eng=eng("D", "", "", f"FORTE-FSCI-RAC-{l}"))
    for x, l in zip("ABC", ("04", "10", "16"))
] + [
    E("DIL", "CI", "Conjunto de Combate a Incêndio por Dilúvio", "RAC-22", 4, sap=["FORTE-FSCI-RAC-22"],
      eng=eng("D", "6 válvulas de dilúvio sem status avaliado, precisam de levantamento", "", "FORTE-FSCI-RAC-22",
              pend=True)),

    # ---------------- Sistema Elétrico ----------------
    E("TFA1", "EL", "Transformador MC A-01", "TF-UC-123101A-01", 3,
      eng=eng("I", "Ensaios elétricos e instalação de para-raio", "Criar OS", "TF-UC-123101A-01")),
    E("TFA2", "EL", "Transformador MC A-02", "TF-UC-123101A-02", 3,
      eng=eng("I", "Ensaios elétricos e instalação de para-raio", "Criar OS", "TF-UC-123101A-02")),
] + [
    E(f"TF{t}{n}", "EL", f"Transformador MC {t}-0{n}", f"TF-UC-123101{t}-0{n}", 3,
      eng=eng("D", "", "", f"TF-UC-123101{t}-0{n}"))
    for t in "BC" for n in (1, 2)
] + [
    E("UPS1", "EL", "UPS 810001 A/B (P08, painel 514015)", "UPS-810001", 4, sap=["FORTE-FECC-UPS-01"],
      eng=eng("R", "Troca do banco de baterias B (02/27)", "Criar OS", "810001 A/B P08 PAINEL 514015")),
    E("UPS2", "EL", "UPS 810002 A/B (Casario, painel 514023)", "UPS-810002", 4,
      sap=["FORTE-FEIE-PNN-01-04", "FORTE-FEIE-PNN-04-04"],
      eng=eng("D", "", "", "810002 A/B CASARIO PAINEL 514023")),
    E("UPS3", "EL", "UPS 551005 (Sistema PAGA)", "UPS-551005", 4, sap=["FORTE-FETL-RET-04-13"],
      eng=eng("D", "", "", "551005 UPS SISTEMA PAGA")),
    E("UPS4", "EL", "UPS 551001/551002 (Rádio)", "UPS-551001", 3,
      sap=["FORTE-FETL-PNN-01-01", "FORTE-FETL-PNN-01-10", "FORTE-FETL-RET-04-01", "FORTE-FETL-TRF-01"],
      eng=eng("D", "", "", "551001 /551002 RADIO")),
    E("UPS5", "EL", "UPS 551003 A/B (Telefonia)", "UPS-551003", 2,
      sap=["FORTE-FETL-PNN-01-07", "FORTE-FETL-RET-04-07", "FORTE-FETL-RET-04-10"],
      eng=eng("D", "", "", "551003 A /B TELEFONIA")),

    # ---------------- Telecomunicações ----------------
    E("PAGA", "TC", "PAGA (Public Address / General Alarm)", "FORTE-FSTI-INT-01", 5, sap=["FORTE-FSTI-INT"],
      eng=eng("R", "Tratamento das anomalias encontradas na inspeção anual em andamento",
              "4355936 / 4382199 / 4176993", "FORTE-FSTI-INT-01")),
    E("PABX", "TC", "PABX", "FORTE-FSTI-TLF", 2, sap=["FORTE-FSTI-TLF"], eng=eng("D", "", "", "FORTE-FSTI-TLF")),
    E("RAD", "TC", "Rádio", "FORTE-FSTI-RAD", 4, sap=["FORTE-FSTI-RAD"], eng=eng("D", "", "", "FORTE-FSTI-RAD")),
    E("GMD", "TC", "GMDSS", "FORTE-FSTI-GMD", 4, sap=["FORTE-FSTI-GMD"], eng=eng("D", "", "", "FORTE-FSTI-GMD")),
    E("CTV", "TC", "CCTV", "FORTE-FSTI-CTV", 2, sap=["FORTE-FSTI-CTV"], eng=eng("D", "", "", "FORTE-FSTI-CTV")),
    E("MET", "TC", "Estação Meteorológica", "FORTE-SDMT", 3, sap=["FORTE-SDMT"], eng=eng("D", "", "", "FORTE-SDMT")),
    E("HEL", "TC", "Monitoramento de Helideck", "FORTE-SDMT", 3, eng=eng("D", "", "", "FORTE-SDMT")),

    # ---------------- Guindastes ----------------
    E("GD1", "GD", "Guindaste de Proa (treliçado, T3303)", "GD-526001", 4, sap=["FORTE-FSMC-GUI-01"],
      oil=["GD-526001"], eng=eng("R", "Inspeção quinquenal em andamento (previsão 11/26)", "", "FORTE-FSMC-GUI-01")),
    E("GD2", "GD", "Guindaste de Meia-Nau (articulado, T3301)", "GD-526002", 3, sap=["FORTE-FSMC-GUI-04"],
      oil=["GD-526002", "GD-526002B02"],
      eng=eng("D", "Inspeção anual prevista para novembro 2026", "4392492", "FORTE-FSMC-GUI-04")),
    E("GD3", "GD", "Guindaste de Popa (articulado, T3302)", "GD-526003", 3, sap=["FORTE-FSMC-GUI-07"],
      oil=["GD-526003"], eng=eng("D", "Inspeção anual prevista para janeiro 2027", "4392493", "FORTE-FSMC-GUI-07")),
    E("MOV", "GD", "Equipamentos de Movimentação de Convés (Z-600004 / Z-644001B)", "Z-600004", 2,
      sap=["FORTE-FSMC-MOA"], oil=["Z-600004", "Z-6440001B"]),
] + [
    # ---------------- Água de Resfriamento ----------------
    E(f"RF{x}", "RF", f"Bomba de Água de Resfriamento {x}", f"B-512401{x}", 4, sap=[f"FORTE-FSRA-MBB-{l}"],
      vib=[f"B-512401{x}", f"M-B-512401{x}"])
    for x, l in zip("ABCD", ("01", "04", "07", "10"))
] + [
    E("RFEA", "RF", "Bomba de Resfriamento de Emergência A", "B-512402A", 3, sap=["FORTE-FSRA-MBB-13"],
      vib=["B-512402A", "MB-512402A"]),
    E("RFEB", "RF", "Bomba de Resfriamento de Emergência B", "B-512402B", 3, sap=["FORTE-FSRA-MBB-16"],
      vib=["B-512402B", "MB-512402B"]),
] + [
    # ---------------- Água Quente ----------------
    E(f"AQ{x}", "AQ", f"Bomba de Água Quente {x}", f"B-512501{x}", 3, sap=[f"FORTE-FSAQ-MBB-{l}"],
      vib=[f"B-512501{x}", f"M-B-512501{x}"])
    for x, l in zip("ABC", ("01", "04", "07"))
] + [
    # ---------------- Escoamento de Petróleo ----------------
    E(f"EP{x}", "EP", f"Bomba de Transferência de Óleo {x}", f"B-122302{x}", 4, sap=[f"FORTE-PSEP-MBB-01-{l}"],
      vib=[f"B-122302{x}", f"M-B-122302{x}"])
    for x, l in zip("ABCDEF", ("01", "04", "07", "10", "13", "16"))
] + [
    # ---------------- Tratamento de Óleo ----------------
    E("SGA", "TO", "Separador de Produção A", "SG-122301A", 5, sap=["FORTE-PSTP-SEP-01"]),
    E("SGB", "TO", "Separador de Produção B", "SG-122301B", 5, sap=["FORTE-PSTP-SEP-04"]),
    E("HTA", "TO", "Aquecedor de Produção A", "P-122303A", 4, sap=["FORTE-PSTP-SEP-01-19"], cases=["20260058"]),
    E("HTB", "TO", "Aquecedor de Produção B", "P-122303B", 4, sap=["FORTE-PSTP-SEP-04-19"]),
] + [
    # ---------------- Água Produzida ----------------
    E(f"AP{x}", "AP", f"Bomba de Água Produzida {x}", f"B-533101{x}", 3, sap=[f"FORTE-PSTA-MBB-01-{l}"],
      vib=[f"MB-533101{x}"] + (["MB-53310C"] if x == "C" else []))
    for x, l in zip("ABCDE", ("01", "04", "07", "10", "13"))
] + [
    E(f"AR{x}", "AP", f"Bomba de Reinjeção de Água {x}", f"B-5331501{x}", 3, sap=[f"FORTE-PSTA-MBB-04-{l}"],
      vib=[f"B-5331501{x}", f"M-B-5331501{x}"])
    for x, l in zip("ABC", ("01", "04", "07"))
] + [
    E("FLA1", "AP", "Bomba de Recirculação do Flotador A-02A", "B-FL-533101A-02A", 2,
      vib=["B-FL-533101A-02A", "M-B-FL-533101A-02A"]),
    E("FLA2", "AP", "Bomba de Recirculação do Flotador A-02B", "B-FL-533101A-02B", 2,
      vib=["B-FL-533101A-02B", "M-B-FL-533101A-02B"]),
    E("FLB1", "AP", "Bomba de Recirculação do Flotador B-02A", "B-FL-533101B-02A", 2,
      vib=["M-B-FL-533101B-02A"]),
    E("FLB2", "AP", "Bomba de Recirculação do Flotador B-02B", "B-FL-533101B-02B", 2,
      vib=["M-B-FL-533101B-02B"]),

    # ---------------- Controle de Poços ----------------
    E("UH1", "CP", "HPU Submarina UH-121001", "UH-121001", 5, sap=["FORTE-PSCP-HPU-01"],
      oil=["TQ-RE-UH-12001-1", "TQ-RE-UH-12001-2", "TQ-RE-UH-121001-1-02", "UH121001", "CIPF2056"]),
    E("UH2", "CP", "HPU de Válvulas UH-121002", "UH-121002", 4, sap=["FORTE-PSCP-HPU-04"],
      oil=["TQ-RE-UH-121002-01", "TQ-RE-UH-121002-02", "UH-121002"]),
] + [
    # ---------------- Ar Comprimido ----------------
    E(f"AC{x}", "AR", f"Compressor de Ar {x}", f"C-UC-513401{x}", 4, sap=[f"FORTE-FSAC-CAR-{l}"])
    for x, l in zip("ABC", ("01", "04", "07"))
] + [
    E("ACD", "AR", "Compressor de Ar Diesel", "C-UC-513402", 3, sap=["FORTE-FSAP-CAR-01"],
      oil=["C-UC-513402", "M-UC-513402", "UC-53402"]),
    E("ACP", "AR", "Compressor de Ar de Partida", "C-UC-513102", 3, oil=["C-UC-513102"]),

    # ---------------- HVAC ----------------
    E("CHI", "HV", "Chillers UR-525201 A/B/C", "UR-525201", 3, sap=["FORTE-FSAR-CHI"],
      vib=[f"UR 5201 {x}{n}" for x in "ABC" for n in (1, 2)]),
    E("AGA", "HV", "Bomba de Água Gelada B-525201A", "B-525201A", 2, sap=["FORTE-FSAR-MBB-01"],
      vib=["B-525201A", "M-525201A"]),
    E("AGB", "HV", "Bomba de Água Gelada B-525201B", "B-525201B", 2, sap=["FORTE-FSAR-MBB-04"],
      vib=["B-525201B", "M-525201B"]),
    E("AGC", "HV", "Bomba de Água Gelada B-525201C", "B-525201C", 2, sap=["FORTE-FSAR-MBB-07"],
      vib=["B-525201C", "MB-525201C"]),
] + [
    E(f"AG{y}", "HV", f"Bomba de Água Gelada B-525202{x}", f"B-525202{x}", 2, sap=[f"FORTE-FSAR-MBB-{l}"],
      vib=[f"B-525202{x}", f"MB-525202{x}"])
    for x, y, l in zip("ABC", "DEF", ("10", "13", "16"))
] + [
    # ---------------- Unidades Hidráulicas / Marinha ----------------
    E("U3A", "MA", "UH de Carga e Lastro UH-600003A", "UH-600003A", 3, sap=["FORTE-PSOF-HPU-01-04"],
      vib=["M-UM-60003A-01A", "M-UM-60003A-01B", "M-UM-60003A-01C"], oil=["UH-600003A"]),
    E("U3B", "MA", "UH de Carga e Lastro UH-600003B", "UH-600003B", 3, sap=["FORTE-PSOF-HPU-01-07"],
      vib=["M-UM-60003B-01A", "M-UM-60003B-01B", "M-UM-60003B-01C"], oil=["UH-600003B"]),
    E("U4", "MA", "UH de Carga e Lastro UH-600004", "UH-600004", 3, sap=["FORTE-PSOF-HPU-01-10"],
      vib=["M-UM-60004-01A", "M-UM-60004-01B"], oil=["UH-600004"]),
    E("U661", "MA", "HPU de Offloading UH-661001", "UH-661001", 3,
      vib=["M-BUH661001 A", "M-BUH661001 B1", "M-BUH661001 B2"],
      oil=["UH-661001", "TP-UH-661004", "TQ-UH-661004-A", "UH-661004-AB"]),
    E("U665", "MA", "HPU de Amarração UH-665101", "UH-665101", 2, oil=["UH-665101", "TQ-UH-665101-B", "UH-665103"]),
    E("U5331", "MA", "UH Casa de Bombas UH-533501", "UH-533501", 3,
      vib=["UH-533501 A", "UH-533501 B", "UH-53501 C"], oil=["TQ-UH-533501"]),
    E("PULL", "MA", "Guincho de Pull-in GN-610001 e HPU", "GN-610001", 2,
      oil=["GN-610001-02", "GN-UH-610001", "UH-610001", "PN-EH-610001-02"]),
    E("U5332", "MA", "UH Praça de Máquinas UH-533502", "UH-533502", 2, oil=["TQ-UH-533502", "UH-533502"]),
]

EQUIPMENT += [
    # ---------------- Salvatagem ----------------
    E("BOT", "SV", "Bote de Resgate EB-543001", "EB-543001", 4, sap=["FORTE-FSSV-BTE-01"],
      oil=["EB-543001", "CIP-F2024"]),
] + [
    E(f"BAL{x}", "SV", f"Baleeira EB-543002{x}", f"EB-543002{x}", 4, sap=[f"FORTE-FSSV-BEE-{l}"],
      oil=[f"TN-EB-54{'2' if x == 'A' else '3'}002{x}"])
    for x, l in zip("ABCD", ("04", "10", "01", "07"))
] + [
    # ---------------- Óleo Diesel ----------------
    E("DSA", "OD", "Bomba de Distribuição de Diesel A", "B-513301A", 3, sap=["FORTE-FSOD-MBB-01-01"],
      vib=["MB-513301 A"]),
    E("DSB", "OD", "Bomba de Distribuição de Diesel B", "B-513301B", 3, sap=["FORTE-FSOD-MBB-01-04"],
      vib=["MB-513301-B"]),
    E("DS4", "OD", "Bomba de Diesel B-513304", "B-513304", 2, sap=["FORTE-FSOD-MBB-01-10"],
      vib=["B- 513304", "MB- 513304"]),
    E("DSI", "OD", "Bomba de Injeção de Diesel nos Poços", "B-513303", 3, sap=["FORTE-FSOD-MBB-04"],
      vib=["M-B 513303"], oil=["B-513303"]),
    E("CEN", "OD", "Centrífugas de Diesel SC-513301A/B", "SC-513301", 2, sap=["FORTE-FSOD-CEN-01"]),
]

# Ativos da lista de alarmes (alarms-list.csv, coluna "Asset Name" sem o prefixo FORTE-)
# associados a cada equipamento.
ALARM_ASSETS = {
    "TGA": ["TS-GE-514001A", "GE-514001A"],
    "TGB": ["TS-GE-514001B", "GE-514001B"],
    "TGC": ["TS-GE-514001C", "GE-514001C"],
    "TGD": ["TS-GE-514001D", "GE-514001D"],
    "MCA": ["C-UC-123101A-01", "C-UC-123101A-02", "VV-C-UC-123101A", "M-UC-123101A"],
    "MCB": ["C-UC-123101B-01", "C-UC-123101B-02", "VV-C-UC-123101B", "M-UC-123101B"],
    "MCC": ["C-UC-123101C-01", "C-UC-123101C-02", "VV-C-UC-123101C", "M-UC-123101C"],
    "BOO": ["UC-1225"],
    "BIA": ["B-UB-125101A"], "BIB": ["B-UB-125101B"], "BIC": ["B-UB-125101C"], "BID": ["B-UB-125101D"],
    "CAA": ["B-511101A"], "CAB": ["B-511101B"], "CAC": ["B-511101C"], "CAD": ["B-511101D"],
    "HTA": ["P-122303A"], "HTB": ["P-122303B"],
    "DGSC1": ["DGS-123101C"], "DGSC2": ["DGS-123101C"],
}

# Cases transcritos das telas "Cases Forte" e "Cases Forte - pt2".
# sev = barras de criticidade preenchidas na 1ª coluna (1 baixa, 2 média, 3 alta, 4 crítica).
CASES = [
    dict(id="20260064", dt="2026-10-08", eq="FORTE-TS-GE-514001D (Turbina a Gás)", title="FM - System 1 Event - Perda/erro de leitura", tipo="Rotativos", st="No prazo", al=0, sev=1, ed="2026-10-08"),
    dict(id="20260061", dt="2026-10-02", eq="FORTE-TS-GE-514001C (Turbina a Gás)", title="FM - System 1 Event - Sensor Axial PT", tipo="Rotativos", st="No prazo", al=21, sev=2, ed="2026-10-02"),
    dict(id="20260059", dt="2026-09-24", eq="FORTE-C-UC-123101B-01 (LP/IP Compressor Trem B)", title="Vazamento de Óleo Lubrificante", tipo="Rotativos", st="No prazo", al=15, sev=3, ed="2026-10-06"),
    dict(id="20260058", dt="2026-09-18", eq="FORTE-P-122303A (Aquecedor Produção A)", title="FM - Perda de Performance - Alta Temperatura na Jusante", tipo="Processo", st="Atrasado", al=50, sev=3, ed="2026-10-14"),
    dict(id="20260052", dt="2026-09-04", eq="FORTE-B-UB-125101B (Bomba de Injeção)", title="Flatline - Temperatura - Mancal da Bomba LOA", tipo="Instrumentação", st="No prazo", al=0, sev=3, ed="2026-09-04"),
    dict(id="20260050", dt="2026-09-01", eq="FORTE-B-UB-125101C (Bomba de Injeção)", title="Falha sensor TT-1251209C Temperatura de Mancal", tipo="Instrumentação", st="No prazo", al=0, sev=1, ed="2026-09-30"),
    dict(id="20260049", dt="2026-08-24", eq="FORTE-C-UC-123101B-02 (HP Compressor Trem B)", title="Anomalia no sistema de óleo lubrificante (Superaquecimento do mancal)", tipo="Rotativos", st="No prazo", al=50, sev=3, ed="2026-08-26"),
    dict(id="20260047", dt="2026-08-20", eq="FORTE-B-UB-125101D (Bomba de Injeção)", title="FM - Alta Vibração no Mancal", tipo="Rotativos", st="No prazo", al=179, sev=3, ed="2026-10-06"),
    dict(id="20260043", dt="2026-08-05", eq="FPSO Forte", title="FPSO Forte - Análise de MPC e RPVOT", tipo="Rotativos", st="No prazo", al=0, sev=2, ed="2026-09-15"),
    dict(id="20260038", dt="2026-07-23", eq="FORTE-C-UC-123101C-01 (LP/IP Compressor Trem C)", title="FM - Alta Vibração no Mancal", tipo="Rotativos", st="Todos Executados", al=46, sev=2, ed="2026-10-06"),
    dict(id="20260035", dt="2026-06-23", eq="FORTE-B-UB-125101C (Bomba de Injeção)", title="FM - Alta Vibração no Mancal", tipo="Rotativos", st="Todos Executados", al=22, sev=3, ed="2026-10-06"),
    dict(id="20260027", dt="2026-05-06", eq="FORTE-TS-GE-514001A (Turbina a Gás)", title="FM - Alto Deslocamento no Mancal Axial", tipo="Instrumentação", st="Atrasado", al=0, sev=3, ed="2026-09-04"),
    dict(id="20260026", dt="2026-04-27", eq="FORTE-UC-1225 (Booster Compressor)", title="FC - Alta Temperatura do Óleo Lubrificante", tipo="Rotativos", st="Atrasado", al=7, sev=4, ed="2026-09-08"),
    dict(id="20260024", dt="2026-04-17", eq="FORTE-GE-514001C (Gerador Elétrico)", title="FORTE - Trip por alta temperatura no sensor da Gearbox - GTG C", tipo="Rotativos", st="No prazo", al=0, sev=2, ed="2026-08-31"),
    dict(id="20250040", dt="2025-12-19", eq="FORTE-C-UC-123101C-02 (HP Compressor Trem C)", title="FM - Alta Vibração no Mancal", tipo="Rotativos", st="Atrasado", al=3, sev=4, ed="2026-08-18"),
    dict(id="20250039", dt="2025-11-21", eq="FORTE-TS-GE-514001C (Turbina a Gás)", title="FC - Lube Oil Low Pressure", tipo="Instrumentação", st="Atrasado", al=0, sev=1, ed="2026-07-07"),
    dict(id="20250025", dt="2025-10-24", eq="FORTE-DGS-123101C (Compressor Alternativo)", title="Compressores N2 DGS - Sensores com leitura anômala", tipo="Instrumentação", st="Todos Executados", al=0, sev=1, ed="2026-08-07"),
    dict(id="20250019", dt="2025-10-08", eq="FORTE-TS-GE-514001D (Turbina a Gás)", title="FM - Axial Bearing High Displacement", tipo="Rotativos", st="No prazo", al=0, sev=3, ed="2026-09-15"),
]
