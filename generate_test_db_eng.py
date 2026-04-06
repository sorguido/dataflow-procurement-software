#!/usr/bin/env python3
"""
Script per generare un database di test completo per DataFlow.

Genera:
  - 500 RFQ distribuite negli anni 2020-2026 (~71-72/anno)
  - 3 fornitori per RFQ con prezzi in offerte_ricevute
  - 100 eventi Saving + 100 Cost Avoidance + 100 Derisking
  - Distribuzione temporale 2020-2026

Database di output: test_dataflow_full.db (NON sovrascrive DB reali).

Strategia numeri:
  - offerte_ricevute.prezzo_unitario: VARCHAR con virgola decimale ("123,4567")
  - vsm_events (REAL): Python float puro (es. 12500.75) — regola virgola e' solo UI layer
"""

import os
import random
from datetime import datetime, date, timedelta

from database_manager import DatabaseManager
from models.vsm_event import VSMEvent
from services.vsm_persistence import save_event_with_impacts

# ---------------------------------------------------------------------------
# CONFIGURAZIONE
# ---------------------------------------------------------------------------
DB_PATH = 'test_dataflow_full.db'
USERNAME = 'gsoraru'
BUYER_NAME = 'Guido Sorarù'
RANDOM_SEED = 42
random.seed(RANDOM_SEED)

ANNI = list(range(2020, 2027))   # 2020 … 2026
RFQ_TOTALI = 500
VSM_PER_TIPO = 100               # 100 Saving, 100 Cost Avoidance, 100 Derisking

# Fornitori fittizi
SUPPLIERS = [
    'MetalWorks GmbH',
    'Precision Steel Ltd',
    'Industrial Components SRL'
]

# Materiali metalmeccanici tipici (codice, descrizione, tipo)
MATERIALS = [
    ('BRG-001', 'Ball Bearing 6205-2RS', 'full_supply'),
    ('SHF-120', 'Stainless Steel Shaft Ø20x300mm', 'full_supply'),
    ('PLT-304', 'Steel Plate 304SS 10x500x1000mm', 'full_supply'),
    ('GSK-NBR', 'NBR O-Ring Gasket Set', 'full_supply'),
    ('BLT-M12', 'Hex Bolt M12x80 DIN 933 8.8', 'full_supply'),
    ('SPG-C50', 'Compression Spring C50 Steel', 'full_supply'),
    ('WSH-M16', 'Flat Washer M16 DIN 125A', 'full_supply'),
    ('PIN-8x50', 'Dowel Pin ISO 2338 8x50mm', 'full_supply'),
    ('FLG-DN80', 'Flange DN80 PN16 Carbon Steel', 'full_supply'),
    ('CST-AL6061', 'Aluminum Casting Al6061-T6', 'work_order'),
    ('MCH-SS316', 'CNC Machined Part SS316L', 'work_order'),
    ('WLD-FR100', 'Welded Frame Structure', 'work_order'),
    ('TRN-45#', 'Turned Component C45 Steel', 'work_order'),
    ('MIL-4140', 'Milled Housing 4140 Steel', 'work_order'),
    ('GRD-SK3', 'Ground Precision Pin SK3', 'work_order'),
    ('BND-16GA', 'Sheet Metal Bending 16GA', 'work_order'),
    ('DRL-PLAT', 'Drilled Mounting Plate', 'work_order'),
    ('THD-M20', 'Threaded Stud M20x2.5', 'work_order'),
    ('HRD-HRC55', 'Heat Treated Component HRC55', 'work_order'),
    ('ASM-MOD5', 'Assembly Module Type 5', 'work_order'),
    ('BLK-STEEL', 'Steel Blank 100x100x200mm', 'full_supply'),
    ('ROD-BRS', 'Brass Rod Ø30x500mm', 'full_supply'),
    ('TUB-ST37', 'Steel Tube ST37 Ø48x3mm', 'full_supply'),
    ('CHN-08B', 'Roller Chain 08B-1 Simplex', 'full_supply'),
    ('GEA-MOD3', 'Spur Gear Module 3 Z25', 'work_order'),
    ('PUL-GT3', 'Timing Pulley GT3 3M', 'full_supply'),
    ('CLP-QD25', 'Quick Disconnect Coupling', 'full_supply'),
    ('VLV-PN10', 'Ball Valve 1" PN10 Brass', 'full_supply'),
    ('PMP-2HP', 'Centrifugal Pump 2HP', 'full_supply'),
    ('MTR-3PH', '3-Phase Electric Motor 5.5kW', 'full_supply'),
]

# Progetti di riferimento
PROJECTS = [
    'Project Phoenix - Conveyor System',
    'Project Atlas - Hydraulic Press Upgrade',
    'Project Neptune - Offshore Platform',
    'Project Titan - Heavy Duty Crane',
    'Project Mercury - Automated Assembly Line',
    'Project Apollo - Precision Manufacturing',
    'Project Orion - Industrial Robotics',
    'Project Voyager - Marine Equipment',
    'Project Galileo - Quality Control System',
    'Project Darwin - Process Optimization',
]

# ---------------------------------------------------------------------------
# DATI AGGIUNTIVI PER VSM
# ---------------------------------------------------------------------------

VSM_DESCRIPTIONS_SAVING = [
    'Negotiation of S235 steel sheets — unit price reduction',
    'Renegotiation of C45 drawn bars supply — annual volume',
    'Saving on EN10210 structural tubes — competitive bid with 3 suppliers',
    'A2-70 stainless fasteners cost reduction — order consolidation',
    'Negotiation of DIN 933 grade 8.8 bolts — three-year contract',
    'Saving on Fe360 sheet laser cutting — process efficiency',
    'Renegotiation of 304 stainless steel bending — bending parameters',
    'C45 quench and temper heat treatment cost reduction — minimum batch',
    'Electrolytic zinc plating negotiation — volume increase',
    'Industrial painting RAL 7035 saving — framework agreement',
    'SS316L CNC turning cost reduction — dedicated tooling',
    'PN16 DN100 carbon steel flanges negotiation — two-year supply',
    'SKF 6205-2RS bearings saving — authorized distributor agreement',
    '4140 part milling cost reduction — optimized process',
    'HEA 100 S355 structural sections negotiation — ex-works pickup',
    'NBR gaskets saving — stock consignment agreement',
    'EN-AC-46000 aluminium stamping cost reduction — annual volume',
    'ISO 606 simplex chain negotiation — three-year supply',
    'Module 3 spur gear saving — purchasing cooperative',
    'Compression spring renegotiation — EN 10270 certification',
]

VSM_DESCRIPTIONS_COST_AVOIDANCE = [
    'Avoided steel sheet price revision — fixed-price clause 12 months',
    'Blocked CW614N brass bar price increase — advance agreement',
    'Avoided turning cost increase — semi-annual fixed-price contract',
    'Countered 6061 aluminium raw material cost increase — hedging',
    'Avoided energy surcharge increase on 316L stainless steel',
    'Blocked raw material transport cost increase — logistics agreement',
    'Avoided CNC machining rush surcharge — predictive supply plan',
    'Countered hot-dip galvanizing price rise — advance volume commitment',
    'Avoided copper scarcity surcharge — advance purchase',
    'Blocked epoxy coating price increase — annual contract',
    'Avoided roller bearing price revision — fixed price list agreement',
    'Countered GJL-250 grey cast iron cost increase — annual pre-order',
    'Avoided stamping cost increase — mold sharing agreement',
    'Blocked heat treatment cost rise — reserved capacity',
    'Avoided flange supply rush surcharge — dedicated safety stock',
    'Countered CNC cutting tool consumable cost increase',
    'Avoided treatment furnace energy surcharge — utilities agreement',
    'Blocked CMM inspection cost increase — servicing contract',
    'Avoided stainless steel special screw price revision — annual forecast',
    'Countered industrial packaging cost increase on supplies',
]

VSM_DESCRIPTIONS_DERISKING = [
    'Qualified second supplier for S355 steel sheets — dependency reduction',
    'Introduced alternative supplier for drawn bars — dual sourcing activated',
    'Qualified new subcontractor for CNC turning — process validated',
    'Flange supply diversification — second EN-certified supplier',
    'Introduced backup supplier for heat treatments — production continuity',
    'Qualified second zinc plating supplier — active standby agreement',
    'Reduced concentration on single special fastener supplier',
    'Introduced European alternative to critical Asian supplier',
    'Qualified new aluminium stamper — re-source plan completed',
    'Bearing supply diversification — agreement with 2 distributors',
    'Introduced proximity supplier for critical supply chain components',
    'Qualified second supplier for NBR/FKM special gaskets',
    'Critical lead time reduction via alternative local supplier',
    'Introduced backup supplier for anti-corrosion coating',
    'Qualified new partner for precision grinding operations',
    'Industrial chain supply diversification — multi-sourcing agreement',
    'Introduced second supplier for S355 structural sections',
    'Qualified alternative supplier for module 3 spur gears',
    'Geopolitical risk reduction — re-source from Eastern European to EU supplier',
    'Introduced Italian supplier for grey iron castings — previously single-sourced',
]

VSM_REFERENCES = [
    'RFQ-2020-{:03d}', 'FN-{}-METAL', 'PO-{:04d}', 'NOTE-{:03d}',
    'PROJ-CONV-{}', 'PROJ-HYD-{}', 'PROJ-ROB-{}', 'CONTR-{:04d}',
]

NEW_SUPPLIERS_DERISKING = [
    'Acciai Speciali Valpadana SRL', 'Eurosteel Componenti SpA',
    'MetalTech Brescia SRL', 'Ferriere del Nord SpA',
    'Lavorazioni Meccaniche Bergamo SRL', 'Officine Guidetti SRL',
    'Trattamenti Termici Padova SpA', 'Galvanica Lombarda SRL',
    'Verniciatura Industriale Veneta SRL', 'Fonderia Bresciana SpA',
    'CNC Precision Parts GmbH', 'Nordic Steel Components AB',
    'Meccanica di Piacenza SRL', 'Forgiatura Toscana SpA',
    'Stampaggio Metalli Emilia SRL', 'Costruzioni Meccaniche Piemonte SRL',
]

# ---------------------------------------------------------------------------
# FUNZIONI HELPER
# ---------------------------------------------------------------------------

def random_date_in_year(year: int) -> date:
    """Restituisce una data ISO casuale nell'anno specificato."""
    start = date(year, 1, 1)
    # Per 2026 limitiamo a fine marzo (data corrente: 2 aprile 2026)
    if year == 2026:
        end = date(2026, 3, 31)
    else:
        end = date(year, 12, 31)
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def generate_price(min_price=5.0, max_price=500.0) -> str:
    """Prezzo con virgola decimale a 4 cifre — formato VARCHAR richiesto da DataFlow."""
    price = random.uniform(min_price, max_price)
    return f"{price:.4f}".replace('.', ',')


def generate_quantity() -> str:
    """Quantità come stringa intera."""
    return str(random.choice([1, 2, 5, 10, 20, 25, 50, 100, 200, 500, 1000]))



# ---------------------------------------------------------------------------
# GENERAZIONE DATABASE
# ---------------------------------------------------------------------------

def create_test_database():
    """Genera il database di test completo: 500 RFQ + 300 eventi VSM."""
    # --- 0. Pulizia ---
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"Database precedente rimosso: {DB_PATH}")

    # --- 1. Schema via DatabaseManager ---
    print("Creazione schema ...")
    db = DatabaseManager(DB_PATH)
    db.create_tables()
    conn = db.conn

    # --- 2. RFQ (500) ---
    print(f"\nGenerazione {RFQ_TOTALI} RFQ ...")
    rfq_per_anno = RFQ_TOTALI // len(ANNI)
    rfq_extra    = RFQ_TOTALI % len(ANNI)

    totale_rfq = 0

    conn.execute("BEGIN")
    for idx_anno, anno in enumerate(ANNI):
        n = rfq_per_anno + (1 if idx_anno < rfq_extra else 0)
        base_id = (anno % 100) * 100000  # 2020 → 2000000, 2026 → 2600000

        for i in range(n):
            rfq_id = base_id + i + 1

            em_date = random_date_in_year(anno)
            ex_date = em_date + timedelta(days=random.randint(15, 45))
            emissione_str = em_date.strftime('%Y-%m-%d')
            scadenza_str  = ex_date.strftime('%Y-%m-%d')

            num_mat  = random.randint(1, 3)
            selected = random.sample(MATERIALS, num_mat)
            has_wo   = any(m[2] == 'work_order' for m in selected)
            tipo_rdo = 'Conto lavoro' if has_wo else 'Fornitura piena'
            stato    = 'attiva' if random.random() < 0.8 else 'archiviata'
            riferimento = random.choice(PROJECTS)

            conn.execute(
                "INSERT INTO richieste_offerta "
                "(id_richiesta, data_emissione, data_scadenza, riferimento, note_generali, "
                " stato, tipo_rdo, username) VALUES (?,?,?,?,?,?,?,?)",
                (rfq_id, emissione_str, scadenza_str, riferimento,
                 f"Test RFQ {rfq_id}", stato, tipo_rdo, USERNAME)
            )

            for sup in SUPPLIERS:
                conn.execute(
                    "INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore) VALUES (?,?)",
                    (rfq_id, sup)
                )

            for mat_code, mat_desc, mat_type in selected:
                qty           = generate_quantity()
                codice_grezzo = f"{mat_code}-RAW"     if mat_type == 'work_order' else ''
                dis_grezzo    = f"DWG-{mat_code}-R01" if mat_type == 'work_order' else ''
                mat_cl        = "Grezzo C45"           if mat_type == 'work_order' else ''

                cur = conn.execute(
                    "INSERT INTO dettagli_richiesta "
                    "(id_richiesta, codice_materiale, descrizione_materiale, quantita, "
                    " disegno, codice_grezzo, disegno_grezzo, materiale_conto_lavoro) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (rfq_id, mat_code, mat_desc, qty, f"DWG-{mat_code}",
                     codice_grezzo, dis_grezzo, mat_cl)
                )
                detail_id = cur.lastrowid

                for sup in SUPPLIERS:
                    prezzo = generate_price(10, 400)
                    conn.execute(
                        "INSERT INTO offerte_ricevute "
                        "(id_dettaglio, nome_fornitore, prezzo_unitario) VALUES (?,?,?)",
                        (detail_id, sup, prezzo)
                    )

            totale_rfq += 1

        print(f"  {anno}: {n} RFQ inserite (ID {base_id + 1} ... {base_id + n})")

    conn.execute("COMMIT")
    print(f"  Totale RFQ inserite: {totale_rfq}")

    # --- 3. VSM Events ---
    _genera_vsm(db)

    # --- 4. Potential Suppliers (Derisking tab) ---
    _genera_potential_suppliers(db)

    # --- 5. Verifica conteggi ---
    _verifica_conteggi(db)

    db.close()
    print(f"\n\u2705  Database di test pronto: {DB_PATH}")


def _genera_vsm(db: DatabaseManager):
    """Inserisce 100 Saving + 100 Cost Avoidance + 100 Derisking."""
    print(f"\nGenerazione {VSM_PER_TIPO * 3} eventi VSM ...")

    # Distribuisce gli eventi uniformemente negli anni
    anni_ext = (ANNI * (VSM_PER_TIPO // len(ANNI) + 1))[:VSM_PER_TIPO]

    # ---- Saving ----
    print(f"  Saving ({VSM_PER_TIPO}) ...")
    for i in range(VSM_PER_TIPO):
        anno    = anni_ext[i]
        ev_date = random_date_in_year(anno)
        desc    = random.choice(VSM_DESCRIPTIONS_SAVING)
        ref     = f"RFQ-{anno}-{random.randint(1, 72):03d}"

        if random.random() < 0.80:
            bdg       = round(random.uniform(5_000, 150_000), 2)
            neg       = round(bdg * random.uniform(0.70, 0.95), 2)
            qta_annua = round(random.uniform(1.0, 100.0), 1)
            opex_rip  = random.choice([True, False])
            event = VSMEvent(
                event_date=datetime(anno, ev_date.month, ev_date.day),
                username=USERNAME, buyer=BUYER_NAME,
                event_type='Saving', action='Negoziazione',
                description=desc, reference=ref,
                driver='Prezzo',
                importo_bdg=bdg, importo_negoziato=neg,
                quantita_annua=qta_annua, percent_realizzo=100.0,
                opex_ripetitivo=opex_rip,
            )
        else:
            spending = round(random.uniform(50_000, 500_000), 2)
            gg_att   = random.choice([30, 60, 90])
            gg_neg   = gg_att + random.choice([30, 60, 90])
            event = VSMEvent(
                event_date=datetime(anno, ev_date.month, ev_date.day),
                username=USERNAME, buyer=BUYER_NAME,
                event_type='Saving', action='Negoziazione',
                description=desc, reference=ref,
                driver='Pagamenti',
                spending_annuo=spending,
                giorni_pagamento_attuali=gg_att,
                giorni_pagamento_negoziati=gg_neg,
            )

        save_event_with_impacts(db, event)

    # ---- Cost Avoidance ----
    print(f"  Cost Avoidance ({VSM_PER_TIPO}) ...")
    for i in range(VSM_PER_TIPO):
        anno    = anni_ext[i]
        ev_date = random_date_in_year(anno)
        desc    = random.choice(VSM_DESCRIPTIONS_COST_AVOIDANCE)
        ref     = f"RFQ-{anno}-{random.randint(1, 72):03d}"

        ric     = round(random.uniform(5_000, 200_000), 2)
        neg     = round(ric * random.uniform(0.75, 0.98), 2)
        qta     = round(random.uniform(1.0, 50.0), 1)

        event = VSMEvent(
            event_date=datetime(anno, ev_date.month, ev_date.day),
            username=USERNAME, buyer=BUYER_NAME,
            event_type='Cost Avoidance', action='Negoziazione',
            description=desc, reference=ref,
            driver='Prezzo',
            importo_richiesto_iniziale=ric,
            importo_negoziato=neg,
            quantita_annua=qta, percent_realizzo=100.0,
        )
        save_event_with_impacts(db, event)

    # ---- Derisking ----
    print(f"  Derisking ({VSM_PER_TIPO}) ...")
    for i in range(VSM_PER_TIPO):
        anno    = anni_ext[i]
        ev_date = random_date_in_year(anno)
        desc    = random.choice(VSM_DESCRIPTIONS_DERISKING)
        new_sup = random.choice(NEW_SUPPLIERS_DERISKING)
        ref     = f"DERISKING-{anno}-{i + 1:03d}"

        event = VSMEvent(
            event_date=datetime(anno, ev_date.month, ev_date.day),
            username=USERNAME, buyer=BUYER_NAME,
            event_type='Derisking', action='Derisking',
            description=desc, reference=ref,
            new_supplier=new_sup,
        )
        save_event_with_impacts(db, event)

    print(f"  VSM inseriti: {VSM_PER_TIPO * 3}")


POTENTIAL_SUPPLIERS_DATA = [
    (
        'Acciai Speciali Valpadana SRL', 'Steel', 'Qualificato',
        'Marco Ferretti', 'm.ferretti@acciaivalpadana.it', '+39 0444 512300',
        'www.acciaivalpadana.it',
        'Second qualified supplier for S355 steel sheets. Audit passed 2024-11. EN 10025 certification.',
        '2024-11-15', '2024-11-15T10:30:00',
    ),
    (
        'Eurosteel Componenti SpA', 'Steel', 'In valutazione',
        'Laura Bianchi', 'l.bianchi@eurosteel.eu', '+39 030 7742100',
        'www.eurosteel.eu',
        'Alternative supplier for C45 drawn bars. Sample qualification in progress. Offer received 2025-02.',
        '2025-02-10', '2025-02-10T09:00:00',
    ),
    (
        'MetalTech Brescia SRL', 'Mechanical machining', 'Qualificato',
        'Giovanni Rossi', 'g.rossi@metaltech-bs.it', '+39 030 3451200',
        'www.metaltech-bs.it',
        'Qualified subcontractor for CNC turning on SS316L. Production capacity validated. Lead time 10 days.',
        '2024-09-20', '2024-09-20T14:00:00',
    ),
    (
        'Nordic Steel Components AB', 'Steel', 'Nuovo',
        'Erik Lindstrom', 'e.lindstrom@nordicsteel.se', '+46 31 7001200',
        'www.nordicsteel.se',
        'Swedish supplier identified as alternative to Eastern European supplier. First commercial visit 2025-03.',
        '2025-03-05', '2025-03-05T11:00:00',
    ),
    (
        'Trattamenti Termici Padova SpA', 'Heat treatments', 'Qualificato',
        'Stefano Meneghetti', 's.meneghetti@ttpd.it', '+39 049 8823400',
        'www.ttpd.it',
        'Qualified backup for C45 hardening and case hardening. Active standby agreement. Capacity 2 t/day.',
        '2024-06-01', '2024-06-01T08:30:00',
    ),
    (
        'Galvanica Lombarda SRL', 'Surface treatments', 'In valutazione',
        'Paola Carminati', 'p.carminati@galvanicalombarda.it', '+39 02 9290500',
        'www.galvanicalombarda.it',
        'Second supplier for electrolytic zinc plating and nickel plating. Samples sent 2025-01. Awaiting quality report.',
        '2025-01-20', '2025-01-20T10:00:00',
    ),
    (
        'Verniciatura Industriale Veneta SRL', 'Surface treatments', 'Qualificato',
        'Antonio Zilio', 'a.zilio@viv-srl.it', '+39 0422 631800',
        'www.viv-srl.it',
        'Backup supplier for epoxy coating RAL 7035. Homologation completed 2023-10. Annual framework agreement.',
        '2023-10-12', '2023-10-12T15:00:00',
    ),
    (
        'Fonderia Bresciana SpA', 'Castings', 'In valutazione',
        'Franco Pezzotti', 'f.pezzotti@fonderiabresciana.it', '+39 030 9821000',
        'www.fonderiabresciana.it',
        'Italian alternative for GJL-250 grey cast iron castings. First sampling in progress. Expected by 2025-05.',
        '2025-02-28', '2025-02-28T09:30:00',
    ),
    (
        'CNC Precision Parts GmbH', 'Mechanical machining', 'Qualificato',
        'Klaus Weber', 'k.weber@cnc-precision.de', '+49 711 4503200',
        'www.cnc-precision.de',
        'German partner for 4140 milling and precision grinding. ISO 9001:2015. Lead time 15 days. Active backup since 2023.',
        '2023-07-18', '2023-07-18T13:00:00',
    ),
    (
        'Lavorazioni Meccaniche Bergamo SRL', 'Mechanical machining', 'Scartato',
        'Roberto Cattaneo', 'r.cattaneo@lmbergamo.it', '+39 035 3310800',
        '',
        'Failed quality audit 2024-04. Out-of-spec tolerances on turned samples. Reassess after corrective actions.',
        '2024-04-10', '2024-04-10T16:00:00',
    ),
    (
        'Stampaggio Metalli Emilia SRL', 'Stamping', 'Qualificato',
        'Elena Monti', 'e.monti@sme-srl.it', '+39 059 8812200',
        'www.sme-srl.it',
        'Second supplier for EN-AC-46000 aluminium stamping. Re-source completed 2024-03. Guaranteed annual volume 4,000 pcs.',
        '2024-03-22', '2024-03-22T10:30:00',
    ),
    (
        'Meccanica di Piacenza SRL', 'Mechanical machining', 'In valutazione',
        'Andrea Ferri', 'a.ferri@meccanicapc.it', '+39 0523 601500',
        'www.meccanicapc.it',
        'Local supplier for lead time reduction on critical parts. Factory visit 2025-03-18. Adequate CNC capacity.',
        '2025-03-18', '2025-03-18T11:00:00',
    ),
    (
        'Forgiatura Toscana SpA', 'Forging', 'Nuovo',
        'Massimo Landi', 'm.landi@forgiaturatoscana.it', '+39 055 8923100',
        'www.forgiaturatoscana.it',
        'Identified as potential alternative for PN16 forged flanges. Technical documentation requested 2025-04.',
        '2025-04-02', '2025-04-02T09:00:00',
    ),
    (
        'Officine Guidetti SRL', 'Mechanical machining', 'Qualificato',
        'Luca Guidetti', 'l.guidetti@officineguidetti.it', '+39 0372 421600',
        'www.officineguidetti.it',
        'Subcontractor for SK3 precision grinding. Active backup on critical components. EN ISO 9001 certified.',
        '2023-11-05', '2023-11-05T14:00:00',
    ),
    (
        'Costruzioni Meccaniche Piemonte SRL', 'Welded structures', 'In valutazione',
        'Davide Gallo', 'd.gallo@cmp-srl.it', '+39 011 9043200',
        'www.cmp-srl.it',
        'Second supplier for welded frames and EN1090 structures. Welder qualification in progress. Expected by 2025-06.',
        '2025-01-30', '2025-01-30T10:00:00',
    ),
    (
        'Catene Industriali Veneto SRL', 'Mechanical transmissions', 'Qualificato',
        'Mirko Trevisan', 'm.trevisan@cateneveneto.it', '+39 049 9213400',
        'www.cateneveneto.it',
        'Second authorized distributor for ISO 606 08B simplex chains. Active multi-sourcing agreement 2024. Local stock 200 m.',
        '2024-05-14', '2024-05-14T09:00:00',
    ),
    (
        'Ingranaggi Precisi Lombardia SRL', 'Mechanical transmissions', 'Nuovo',
        'Chiara Colombo', 'c.colombo@ingranaggeriprecisi.it', '+39 02 4001500',
        '',
        'Initial contact for module 3 cylindrical gears. Quote request sent 2025-03-28.',
        '2025-03-28', '2025-03-28T15:30:00',
    ),
    (
        'Guarnizioni & Tenute SRL', 'Rubber and seals', 'Qualificato',
        'Fabio Negri', 'f.negri@guarnizionientenute.it', '+39 035 4124500',
        'www.guarnizionientenute.it',
        'Second qualified supplier for NBR and FKM gaskets. Samples approved 2024-08. Annual agreement signed.',
        '2024-08-05', '2024-08-05T11:00:00',
    ),
    (
        'Cuscinetti Nord Italia Srl', 'Bearings and transmissions', 'Qualificato',
        'Simona Riva', 's.riva@cni-srl.it', '+39 02 6682100',
        'www.cni-srl.it',
        'Second authorized SKF distributor for 6205-2RS and roller bearings. Fixed SKF price list 2025. Milan warehouse.',
        '2024-10-01', '2024-10-01T08:30:00',
    ),
    (
        'Profilati Strutturali Sud SpA', 'Steel', 'Scartato',
        'Giuseppe Marino', 'g.marino@profilatisudspa.it', '+39 081 7530900',
        'www.profilatisudspa.it',
        'Rejected due to excessive lead time (60 days vs 20 days required) and lack of EN 10210 certifications. Reconsider only if lead time improves.',
        '2024-07-22', '2024-07-22T15:00:00',
    ),
]


def _genera_potential_suppliers(db: DatabaseManager):
    """Inserisce i 20 fornitori potenziali (tab Derisking)."""
    print(f"\nGenerazione {len(POTENTIAL_SUPPLIERS_DATA)} fornitori potenziali (Derisking) ...")
    db.conn.executemany(
        """
        INSERT INTO potential_suppliers
            (supplier_name, category, supplier_status, contact_name,
             email, phone, website, notes, username, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                name, category, status, contact, email, phone,
                website, notes, USERNAME, created_at, updated_at,
            )
            for name, category, status, contact, email, phone,
                website, notes, created_at, updated_at
            in POTENTIAL_SUPPLIERS_DATA
        ],
    )
    db.conn.commit()
    print(f"  Fornitori potenziali inseriti: {len(POTENTIAL_SUPPLIERS_DATA)}")


def _verifica_conteggi(db: DatabaseManager):
    """Stampa i conteggi di verifica del database generato."""
    print("\n--- Verifica conteggi ---")
    c = db.conn

    def q(sql):
        return c.execute(sql).fetchone()[0]

    rfq_tot    = q("SELECT COUNT(*) FROM richieste_offerta")
    sup_tot    = q("SELECT COUNT(*) FROM richiesta_fornitori")
    det_tot    = q("SELECT COUNT(*) FROM dettagli_richiesta")
    off_tot    = q("SELECT COUNT(*) FROM offerte_ricevute")
    vsm_tot    = q("SELECT COUNT(*) FROM vsm_events")
    vsm_sav    = q("SELECT COUNT(*) FROM vsm_events WHERE event_type='Saving'")
    vsm_ca     = q("SELECT COUNT(*) FROM vsm_events WHERE event_type='Cost Avoidance'")
    vsm_der    = q("SELECT COUNT(*) FROM vsm_events WHERE event_type='Derisking'")
    imp_tot    = q("SELECT COUNT(*) FROM vsm_impacts")
    ps_tot     = q("SELECT COUNT(*) FROM potential_suppliers")

    print(f"  richieste_offerta  : {rfq_tot:>6}  (atteso: {RFQ_TOTALI})")
    print(f"  richiesta_fornitori: {sup_tot:>6}  (atteso: {RFQ_TOTALI * 3})")
    print(f"  dettagli_richiesta : {det_tot:>6}  (ca. {RFQ_TOTALI * 2})")
    print(f"  offerte_ricevute   : {off_tot:>6}  (ca. {det_tot * 3})")
    print(f"  vsm_events tot     : {vsm_tot:>6}  (atteso: {VSM_PER_TIPO * 3})")
    print(f"    Saving           : {vsm_sav:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"    Cost Avoidance   : {vsm_ca:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"    Derisking        : {vsm_der:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"  vsm_impacts        : {imp_tot:>6}")
    print(f"  potential_suppliers: {ps_tot:>6}  (atteso: {len(POTENTIAL_SUPPLIERS_DATA)})")

    print("\n  Distribuzione RFQ per anno:")
    rows = c.execute(
        "SELECT substr(data_emissione,1,4) AS anno, COUNT(*) "
        "FROM richieste_offerta GROUP BY anno ORDER BY anno"
    ).fetchall()
    for row in rows:
        print(f"    {row[0]}: {row[1]:>4} RFQ")

    orfani_det = q(
        "SELECT COUNT(*) FROM dettagli_richiesta d "
        "WHERE NOT EXISTS (SELECT 1 FROM richieste_offerta r WHERE r.id_richiesta=d.id_richiesta)"
    )
    orfani_off = q(
        "SELECT COUNT(*) FROM offerte_ricevute o "
        "WHERE NOT EXISTS (SELECT 1 FROM dettagli_richiesta d WHERE d.id_dettaglio=o.id_dettaglio)"
    )
    print(f"\n  Orfani dettagli_richiesta: {orfani_det}  (atteso: 0)")
    print(f"  Orfani offerte_ricevute  : {orfani_off}  (atteso: 0)")

    ok = (rfq_tot == RFQ_TOTALI and vsm_tot == VSM_PER_TIPO * 3
          and ps_tot == len(POTENTIAL_SUPPLIERS_DATA)
          and orfani_det == 0 and orfani_off == 0)
    print(f"\n{'\u2705  Tutti i controlli superati.' if ok else '\u26a0\ufe0f  Attenzione: verificare i conteggi sopra.'}")


if __name__ == '__main__':
    create_test_database()
