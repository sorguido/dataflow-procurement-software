#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script per aggiungere automaticamente le traduzioni mancanti al file .po
usando Google Translate o traduzioni predefinite
"""
import re
import polib

def translate_to_english(italian_text):
    """Traduzioni predefinite per stringhe comuni"""
    
    # Dizionario con traduzioni predefinite
    translations = {
        # Titoli sezioni guida
        "- Allegati Non Trovati": "- Attachments Not Found",
        "- Analisi SQDC": "- SQDC Analysis",
        "- Avanzate": "- Advanced",
        "- Backup": "- Backup",
        "- Benvenuto in DataFlow": "- Welcome to DataFlow",
        "- Database Bloccato": "- Database Locked",
        "- Errori Importazione Excel": "- Excel Import Errors",
        "- Esportazione Excel": "- Excel Export",
        "- Filtri di Ricerca": "- Search Filters",
        "- Gestione Allegati": "- Attachment Management",
        "- Gestione Database": "- Database Management",
        "- Gestione Numeri Ordine (PO)": "- Purchase Order (PO) Number Management",
        "- Importazione da Excel": "- Import from Excel",
        "- Inserimento Manuale degli Articoli": "- Manual Item Entry",
        "- Interfaccia e Pulsanti Principali": "- Interface and Main Buttons",
        "- La Griglia Prezzi": "- Price Grid",
        "- Modifica Dati e Aggiunta Note": "- Edit Data and Add Notes",
        "- Ordinamento delle Colonne": "- Column Sorting",
        "- Prima RdO di Prova": "- First Test RfQ",
        "- Primo Avvio": "- First Launch",
        "- Recupero da Backup": "- Backup Recovery",
        "- Scorciatoie da Tastiera": "- Keyboard Shortcuts",
        
        # Testi licenza
        "1. CONCESSIONE DELLA LICENZA": "1. LICENSE GRANT",
        "2. RESTRIZIONI": "2. RESTRICTIONS",
        "3. ESCLUSIONE DI GARANZIA": "3. DISCLAIMER OF WARRANTY",
        "4. LIMITAZIONE DI RESPONSABILITÀ": "4. LIMITATION OF LIABILITY",
        "5. PROPRIETÀ INTELLETTUALE": "5. INTELLECTUAL PROPERTY",
        "6. TERMINE E RISOLUZIONE": "6. TERM AND TERMINATION",
        "7. LEGGE APPLICABILE": "7. GOVERNING LAW",
        "8. RACCOLTA E UTILIZZO DEI DATI": "8. DATA COLLECTION AND USE",
        
        # Messaggi vari
        "Avviso": "Warning",
        "Non è possibile aggiungere manualmente più voci di quanto previsto dal modello.": "Cannot manually add more entries than provided by the template.",
        "Categoria non valida selezionata.": "Invalid category selected.",
        "Nessuna voce selezionata da eliminare.": "No entry selected for deletion.",
        "Sei sicuro di voler eliminare la riga selezionata?": "Are you sure you want to delete the selected row?",
        "Riga eliminata con successo.": "Row deleted successfully.",
        "Per le RdO Customizzate puoi aggiungere fino a {} articoli.": "For Custom RfQs you can add up to {} items.",
        "Vuoi aggiungere un altro articolo manualmente?": "Do you want to add another item manually?",
        "Per le RdO Customizzate puoi aggiungere solo {} articoli in totale.\nHai già raggiunto il limite.": "For Custom RfQs you can add only {} items in total.\nYou have already reached the limit.",
        "Seleziona un fornitore per aggiungere un allegato.": "Select a supplier to add an attachment.",
        "Seleziona un fornitore per aprire un allegato.": "Select a supplier to open an attachment.",
        "Seleziona un fornitore per scaricare un allegato.": "Select a supplier to download an attachment.",
    }
    
    # Se la traduzione è predefinita, usala
    if italian_text in translations:
        return translations[italian_text]
    
    # Altrimenti, traduzioni automatiche basate su pattern comuni
    text = italian_text
    
    # Pattern comuni italiano -> inglese
    replacements = {
        "Impossibile ": "Unable to ",
        "Errore ": "Error ",
        "Successo": "Success",
        "Attenzione": "Warning",
        "Conferma": "Confirm",
        "Eliminazione": "Deletion",
        "Seleziona ": "Select ",
        "Inserisci ": "Enter ",
        "Salvato con successo": "Saved successfully",
        "Caricamento": "Loading",
        "non trovato": "not found",
        "non trovata": "not found",
    }
    
    for it, en in replacements.items():
        text = text.replace(it, en)
    
    return text


# Leggi il file sorgente
print("Lettura file sorgente...")
with open(r"c:\Users\sorgu\Il mio Drive\Lavoro\App\App Python\DataFlow\DataFlow 2.0.0\Sorgenti\DataFlow 2.0.0.py", 'r', encoding='utf-8') as f:
    source_code = f.read()

# Cerca tutte le stringhe tradotte con _()
pattern_simple = r'_\(\s*"([^"]+)"\s*\)'
pattern_multiline = r'_\(\s*\n\s*"([^"]*(?:\n[^"]*)*?)"\s*\n\s*\)'

all_strings = set()
matches = re.findall(pattern_simple, source_code)
all_strings.update(matches)
matches_multi = re.findall(pattern_multiline, source_code, re.DOTALL)
all_strings.update(matches_multi)

# Carica il file .po
print("Caricamento file .po...")
po_file_path = r"c:\Users\sorgu\Il mio Drive\Lavoro\App\App Python\DataFlow\DataFlow 2.0.0\Sorgenti\locale\en\LC_MESSAGES\dataflow.po"
po = polib.pofile(po_file_path)
existing_msgids = {entry.msgid for entry in po}

# Trova le stringhe mancanti
missing = []
for msg_str in all_strings:
    cleaned = msg_str.replace('\\n', '\n').strip()
    if cleaned and cleaned not in existing_msgids:
        missing.append(cleaned)

print(f"\nTrovate {len(missing)} stringhe mancanti nel file .po")

if missing:
    print("\nAggiunta traduzioni mancanti...")
    added_count = 0
    
    for italian_text in sorted(missing):
        # Traduci il testo
        english_text = translate_to_english(italian_text)
        
        # Crea una nuova entry
        entry = polib.POEntry(
            msgid=italian_text,
            msgstr=english_text
        )
        
        # Aggiungi al file .po
        po.append(entry)
        added_count += 1
        
        # Mostra progresso ogni 10 stringhe
        if added_count % 10 == 0:
            print(f"  Aggiunte {added_count}/{len(missing)} traduzioni...")
    
    # Salva il file .po aggiornato
    print(f"\nSalvataggio file .po con {added_count} nuove traduzioni...")
    po.save(po_file_path)
    
    # Compila il file .mo
    print("Compilazione file .mo...")
    po.save_as_mofile(po_file_path.replace('.po', '.mo'))
    
    print(f"\n✅ Operazione completata! Aggiunte {added_count} traduzioni.")
else:
    print("\n✅ Nessuna traduzione mancante!")

print("\nPrime 20 traduzioni aggiunte:")
for i, text in enumerate(sorted(missing)[:20], 1):
    preview = text.replace('\n', '\\n')[:60]
    translation = translate_to_english(text).replace('\n', '\\n')[:60]
    print(f"{i}. IT: {preview}")
    print(f"   EN: {translation}\n")
