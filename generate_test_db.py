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
    'Negoziazione lamiere acciaio S235 — riduzione prezzo unitario',
    'Rinegoziazione fornitura barre trafilate C45 — volume annuo',
    'Saving su tubi strutturali EN10210 — gara comparativa 3 fornitori',
    'Riduzione costo viteria inox A2-70 — consolidamento ordini',
    'Negoziazione bulloneria DIN 933 classe 8.8 — contratto triennale',
    'Saving taglio laser lamiera Fe360 — efficienza processo',
    'Rinegoziazione piegatura acciaio inox 304 — parametri di piegatura',
    'Riduzione costo trattamento termico bonifica Su C45 — lotto minimo',
    'Negoziazione zincatura elettrolitica — incremento quantitativi',
    'Saving verniciatura industriale RAL 7035 — accordo quadro',
    'Riduzione costo tornitura CNC Su SS316L — attrezzatura dedicata',
    'Negoziazione flange PN16 DN100 acciaio carbone — fornitura biennale',
    'Saving cuscinetti SKF 6205-2RS — accordo distributore ufficiale',
    'Riduzione costo fresatura particolari 4140 — processo ottimizzato',
    'Negoziazione profilati HEA 100 S355 — ritiro franco magazzino',
    'Saving guarnizioni NBR — stock consignment agreement',
    'Riduzione costo stampaggio alluminio EN-AC-46000 — volume annuo',
    'Negoziazione catene simplex ISO606 — fornitura triennale',
    'Saving ingranaggi cilindrici modulo 3 — cooperativa acquisti',
    'Rinegoziazione molle di compressione — certificazione EN 10270',
]

VSM_DESCRIPTIONS_COST_AVOIDANCE = [
    'Evitata revisione prezzi lamiere acciaio — clausola fissa 12 mesi',
    'Bloccato aumento prezzo barre ottone CW614N — accordo preventivo',
    'Evitato incremento costi tornitura — contratto prezzo fisso semestrale',
    'Contrastato aumento costo materia prima alluminio 6061 — hedging',
    'Evitata maggiorazione energy surcharge Su acciaio inox 316L',
    'Bloccato aumento costo trasporto materiale grezzo — accordo logistica',
    'Evitato extra costo urgenza lavorazioni CNC — piano fornitura predittivo',
    'Contrastato rialzo zincatura a caldo — volume garantito anticipato',
    'Evitato surcharge scarsità materia prima rame — acquisto anticipato',
    'Bloccato aumento prezzo verniciatura epossidica — contratto annuale',
    'Evitata revisione prezzi cuscinetti a rulli — accordo listino fisso',
    'Contrastato aumento costo ghisa grigia GJL-250 — pre-ordine annuale',
    'Evitato incremento costo stampaggio — mold sharing agreement',
    'Bloccato rialzo costo trattamenti termici — capacità riservata',
    'Evitato extra urgenza fornitura flange — safety stock dedicato',
    'Contrastato aumento materiali di consumo utensili CNC',
    'Evitato surcharge energia forni trattamento — accordo utilities',
    'Bloccato incremento costo collaudo CMM — contratto servicing',
    'Evitata revisione prezzi viti speciali acciaio inox — forecast annuo',
    'Contrastato aumento costo imballaggi industriali su forniture',
]

VSM_DESCRIPTIONS_DERISKING = [
    'Qualificato secondo fornitore per lamiere acciaio S355 — riduzione dipendenza',
    'Introdotto fornitore alternativo barre trafilate — dual sourcing attivato',
    'Qualificato nuovo terzista per tornitura CNC — processo validato',
    'Diversificazione fornitura flange — secondo fornitore certificato EN',
    'Introdotto fornitore backup per trattamenti termici — continuità produzione',
    'Qualificato secondo fornitore zincatura — accordo stand-by attivo',
    'Riduzione concentrazione Su un solo fornitore di viteria speciale',
    'Introdotto fornitore europeo alternativo a fornitore asiatico critico',
    'Qualificato nuovo stampatore alluminio — piano re-source completato',
    'Diversificazione approvvigionamento cuscinetti — accordo con 2 distributori',
    'Introdotto fornitore di prossimità per componenti critici supply chain',
    'Qualificato secondo fornitore per guarnizioni speciali NBR/FKM',
    'Riduzione lead time critico tramite fornitore locale alternativo',
    'Introdotto fornitore di backup per verniciatura antiruggine',
    'Qualificato nuovo partner per lavorazioni di rettifica di precisione',
    'Diversificazione fornitura catene industriali — accordo multi-sourcing',
    'Introdotto secondo fornitore per profilati strutturali S355',
    'Qualificato fornitore alternativo per ingranaggi cilindrici modulo 3',
    'Riduzione rischio geopolitico — re-source fornitore est Europa a fornitore UE',
    'Introdotto fornitore italiano per fusioni ghisa — before single-sourced',
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

    # --- 4. Verifica conteggi ---
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

    print(f"  richieste_offerta  : {rfq_tot:>6}  (atteso: {RFQ_TOTALI})")
    print(f"  richiesta_fornitori: {sup_tot:>6}  (atteso: {RFQ_TOTALI * 3})")
    print(f"  dettagli_richiesta : {det_tot:>6}  (ca. {RFQ_TOTALI * 2})")
    print(f"  offerte_ricevute   : {off_tot:>6}  (ca. {det_tot * 3})")
    print(f"  vsm_events tot     : {vsm_tot:>6}  (atteso: {VSM_PER_TIPO * 3})")
    print(f"    Saving           : {vsm_sav:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"    Cost Avoidance   : {vsm_ca:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"    Derisking        : {vsm_der:>6}  (atteso: {VSM_PER_TIPO})")
    print(f"  vsm_impacts        : {imp_tot:>6}")

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
          and orfani_det == 0 and orfani_off == 0)
    print(f"\n{'\u2705  Tutti i controlli superati.' if ok else '\u26a0\ufe0f  Attenzione: verificare i conteggi sopra.'}")


if __name__ == '__main__':
    create_test_database()
