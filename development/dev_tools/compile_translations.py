#!/usr/bin/env python3
"""
Script per compilare i file .po in .mo usando polib.
Compila tutte le traduzioni nelle directory locale/*/LC_MESSAGES/
"""

import os
import polib

def _find_project_root(start_dir):
    """Risalita robusta fino alla root che contiene la directory `locale`."""
    current = os.path.abspath(start_dir)
    while True:
        if os.path.isdir(os.path.join(current, 'locale')):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return os.path.abspath(start_dir)


DEV_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = _find_project_root(DEV_TOOLS_DIR)
LOCALE_DIR = os.path.join(PROJECT_ROOT, 'locale')

def compile_translations():
    """Compila tutti i file .po trovati in locale/*/LC_MESSAGES/"""
    
    # Lista delle lingue disponibili
    languages = ['en', 'it']
    
    for lang in languages:
        po_path = os.path.join(LOCALE_DIR, lang, 'LC_MESSAGES', 'dataflow.po')
        mo_path = os.path.join(LOCALE_DIR, lang, 'LC_MESSAGES', 'dataflow.mo')
        
        if not os.path.exists(po_path):
            print(f"⚠️  File .po non trovato: {po_path}")
            continue
        
        try:
            # Carica il file .po
            po = polib.pofile(po_path)
            
            # Salva come .mo
            po.save_as_mofile(mo_path)
            
            print(f"✅ Compilato: {lang}/LC_MESSAGES/dataflow.mo ({len(po)} entries)")
            
        except Exception as e:
            print(f"❌ Errore compilando {lang}: {e}")
    
    print("\n✨ Compilazione completata!")

if __name__ == '__main__':
    compile_translations()
