#!/usr/bin/env python3
"""
Script per generare un database di test per DataFlow
con RfQ metalmeccaniche realistiche in lingua inglese.
"""

import sqlite3
from datetime import datetime, timedelta
import random
import os

# Configurazione
DB_PATH = 'test_dataflow.db'
USERNAME = 'gsoraru'

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

def generate_date(days_ago, variance=30):
    """Genera una data casuale intorno a days_ago giorni fa"""
    base_date = datetime.now() - timedelta(days=days_ago)
    offset = timedelta(days=random.randint(-variance, variance))
    return (base_date + offset).strftime('%Y-%m-%d')

def generate_price(min_price=5, max_price=500):
    """Genera un prezzo con virgola come separatore decimale (4 decimali)"""
    price = random.uniform(min_price, max_price)
    # Formatta con virgola e 4 decimali: es. 123,4567
    return f"{price:.4f}".replace('.', ',')

def generate_quantity():
    """Genera una quantità casuale"""
    return str(random.choice([1, 2, 5, 10, 20, 25, 50, 100, 200, 500, 1000]))

def create_test_database():
    """Crea il database di test con 30 RfQ"""
    
    # Rimuovi database esistente
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"Database esistente rimosso: {DB_PATH}")
    
    # Connetti al database
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Crea le tabelle (schema completo DataFlow)
    print("Creazione tabelle...")
    
    cursor.execute('''
        CREATE TABLE richieste_offerta (
            id_richiesta INTEGER PRIMARY KEY,
            data_emissione VARCHAR,
            data_scadenza VARCHAR,
            riferimento VARCHAR,
            note_generali VARCHAR,
            stato VARCHAR NOT NULL DEFAULT 'attiva',
            numeri_ordine VARCHAR,
            tipo_rdo VARCHAR NOT NULL DEFAULT 'Fornitura piena',
            note_formattate VARCHAR,
            username VARCHAR
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE dettagli_richiesta (
            id_dettaglio INTEGER PRIMARY KEY AUTOINCREMENT,
            id_richiesta INTEGER,
            codice_materiale VARCHAR,
            descrizione_materiale VARCHAR,
            quantita VARCHAR,
            disegno VARCHAR,
            data_consegna_richiesta VARCHAR,
            codice_grezzo VARCHAR,
            disegno_grezzo VARCHAR,
            materiale_conto_lavoro VARCHAR,
            FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta)
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE richiesta_fornitori (
            id_richiesta INTEGER,
            nome_fornitore VARCHAR,
            PRIMARY KEY (id_richiesta, nome_fornitore),
            FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta)
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE offerte_ricevute (
            id_dettaglio INTEGER,
            nome_fornitore VARCHAR,
            prezzo_unitario VARCHAR,
            PRIMARY KEY (id_dettaglio, nome_fornitore),
            FOREIGN KEY (id_dettaglio) REFERENCES dettagli_richiesta (id_dettaglio)
        )
    ''')
    
    print("Generazione dati...")
    
    # Anno corrente per ID (es. 2026 -> 2600000)
    year = int(datetime.now().strftime('%y'))
    base_id = year * 100000
    
    # Genera 30 RfQ
    rfq_count = 30
    for i in range(rfq_count):
        rfq_id = base_id + i + 1
        
        # Seleziona 1-3 materiali casuali
        num_materials = random.randint(1, 3)
        selected_materials = random.sample(MATERIALS, num_materials)
        
        # Determina il tipo RfQ in base ai materiali selezionati
        has_work_order = any(mat[2] == 'work_order' for mat in selected_materials)
        has_full_supply = any(mat[2] == 'full_supply' for mat in selected_materials)
        
        if has_work_order and has_full_supply:
            tipo_rdo = 'Conto lavoro'  # Mixed, usa work order
        elif has_work_order:
            tipo_rdo = 'Conto lavoro'
        else:
            tipo_rdo = 'Fornitura piena'
        
        # Date
        issue_date = generate_date(days_ago=90, variance=60)
        issue_dt = datetime.strptime(issue_date, '%Y-%m-%d')
        expiry_dt = issue_dt + timedelta(days=random.randint(15, 45))
        expiry_date = expiry_dt.strftime('%Y-%m-%d')
        
        # Riferimento progetto
        riferimento = random.choice(PROJECTS)
        
        # Note
        note = f"Request for {num_materials} item(s) - {tipo_rdo}"
        
        # Stato (80% attive, 20% archiviate)
        stato = 'attiva' if random.random() < 0.8 else 'archiviata'
        
        # Inserisci RfQ
        cursor.execute('''
            INSERT INTO richieste_offerta 
            (id_richiesta, data_emissione, data_scadenza, riferimento, note_generali, 
             stato, tipo_rdo, username)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (rfq_id, issue_date, expiry_date, riferimento, note, stato, tipo_rdo, USERNAME))
        
        # Inserisci fornitori (tutti e 3)
        for supplier in SUPPLIERS:
            cursor.execute('''
                INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore)
                VALUES (?, ?)
            ''', (rfq_id, supplier))
        
        # Inserisci materiali
        for mat_code, mat_desc, mat_type in selected_materials:
            qty = generate_quantity()
            
            # Campi specifici per work order
            codice_grezzo = '' if mat_type == 'full_supply' else f"{mat_code}-RAW"
            disegno_grezzo = '' if mat_type == 'full_supply' else f"DWG-{mat_code}-R01"
            materiale_grezzo = '' if mat_type == 'full_supply' else "C45 Steel Blank"
            
            cursor.execute('''
                INSERT INTO dettagli_richiesta 
                (id_richiesta, codice_materiale, descrizione_materiale, quantita,
                 disegno, codice_grezzo, disegno_grezzo, materiale_conto_lavoro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (rfq_id, mat_code, mat_desc, qty, f"DWG-{mat_code}", 
                  codice_grezzo, disegno_grezzo, materiale_grezzo))
            
            # Ottieni l'ID del dettaglio appena inserito
            detail_id = cursor.lastrowid
            
            # Inserisci prezzi per ciascun fornitore
            for supplier in SUPPLIERS:
                # Varia i prezzi tra fornitori
                if supplier == 'MetalWorks GmbH':
                    price = generate_price(10, 300)
                elif supplier == 'Precision Steel Ltd':
                    price = generate_price(15, 350)
                else:  # Industrial Components SRL
                    price = generate_price(12, 320)
                
                cursor.execute('''
                    INSERT INTO offerte_ricevute (id_dettaglio, nome_fornitore, prezzo_unitario)
                    VALUES (?, ?, ?)
                ''', (detail_id, supplier, price))
        
        print(f"  RfQ {rfq_id} creata: {riferimento} ({tipo_rdo}, {stato})")
    
    # Commit e chiudi
    conn.commit()
    conn.close()
    
    print(f"\n✅ Database di test creato: {DB_PATH}")
    print(f"   - {rfq_count} RfQ generate")
    print(f"   - {len(SUPPLIERS)} fornitori")
    print(f"   - Username: {USERNAME}")
    print(f"   - Mix di Full Supply e Work Order")
    print(f"   - Prezzi con separatore virgola (es. 123,4500)")
    print(f"\nPer utilizzare il database:")
    print(f"   1. Rinomina il tuo database attuale (backup)")
    print(f"   2. Copia {DB_PATH} nella posizione corretta")
    print(f"   3. Avvia DataFlow e accedi come utente 'gsoraru'")

if __name__ == '__main__':
    create_test_database()
