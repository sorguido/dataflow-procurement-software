#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script per compilare i file .po in .mo usando polib
"""
import polib
import os

# Percorsi dei file .po
locale_dir = os.path.dirname(os.path.abspath(__file__))
po_files = [
    os.path.join(locale_dir, 'locale', 'it', 'LC_MESSAGES', 'dataflow.po'),
    os.path.join(locale_dir, 'locale', 'en', 'LC_MESSAGES', 'dataflow.po'),
]

print("=== Compilazione file di traduzione ===\n")

for po_file in po_files:
    if not os.path.exists(po_file):
        print(f"AVVISO: File non trovato: {po_file}")
        continue
    
    mo_file = po_file.replace('.po', '.mo')
    
    try:
        print(f"Compilando: {po_file}")
        po = polib.pofile(po_file)
        po.save_as_mofile(mo_file)
        print(f"OK - Salvato: {mo_file}\n")
    except Exception as e:
        print(f"ERRORE compilando {po_file}: {e}\n")

print("=== Compilazione completata ===")
