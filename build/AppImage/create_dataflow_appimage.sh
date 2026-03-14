#!/usr/bin/env bash

set -u

CURRENT_DIR="$(pwd)"
EXPECTED_DIR="$HOME/AppImage"
APPIMAGE_TOOL="$EXPECTED_DIR/appimagetool-x86_64.AppImage"
APPDIR="$EXPECTED_DIR/DataFlow.AppDir"
OUTPUT_APPIMAGE="$EXPECTED_DIR/DataFlow-x86_64.AppImage"

pause_step() {
    echo
    read -rp "Premi INVIO per continuare..."
    echo
}

error_exit() {
    echo
    echo "ERRORE: $1"
    echo
    echo "Operazione interrotta."
    exit 1
}

check_file_exists() {
    local filepath="$1"
    local description="$2"
    if [ ! -e "$filepath" ]; then
        error_exit "$description non trovato: $filepath"
    fi
}

check_command_exists() {
    local cmd="$1"
    if ! command -v "$cmd" >/dev/null 2>&1; then
        error_exit "Comando richiesto non trovato: $cmd"
    fi
}

echo "========================================"
echo " Creazione guidata AppImage di DataFlow "
echo "========================================"
echo

echo "Verifica cartella di esecuzione..."
if [ "$CURRENT_DIR" != "$EXPECTED_DIR" ]; then
    echo
    echo "ATTENZIONE:"
    echo "Questo script deve essere eseguito dalla cartella:"
    echo "$EXPECTED_DIR"
    echo
    echo "Cosa devi fare:"
    echo "1. Crea la cartella $EXPECTED_DIR se non esiste"
    echo "2. Copia dentro quella cartella tutti i file necessari"
    echo "3. Copia dentro anche questo script"
    echo "4. Entra nella cartella con: cd $EXPECTED_DIR"
    echo "5. Riesegui lo script da lì"
    echo
    exit 1
fi

echo "OK: lo script è stato avviato dalla cartella corretta: $EXPECTED_DIR"
pause_step

echo "Verifica prerequisiti di base..."
check_command_exists dpkg-deb
check_command_exists chmod
check_command_exists readlink
echo "OK: comandi di base presenti."
pause_step

echo "Verifica presenza di appimagetool..."
check_file_exists "$APPIMAGE_TOOL" "appimagetool"
chmod +x "$APPIMAGE_TOOL" || error_exit "Impossibile rendere eseguibile appimagetool"
echo "OK: appimagetool trovato in $APPIMAGE_TOOL"
pause_step

echo "STEP 1 - Inserisci nella cartella $EXPECTED_DIR il file .deb di DataFlow."
echo "Quando l'hai fatto, scrivi il nome esatto del file."
read -rp "Nome file .deb: " DEB_FILE

[ -n "$DEB_FILE" ] || error_exit "Nome file .deb vuoto"
DEB_PATH="$EXPECTED_DIR/$DEB_FILE"
check_file_exists "$DEB_PATH" "Pacchetto .deb"
echo "OK: trovato $DEB_PATH"
pause_step

echo "STEP 2 - Inserisci nella cartella $EXPECTED_DIR l'icona PNG di DataFlow."
echo "Consigliata 256x256, ma va bene anche 150x150."
echo "Quando l'hai fatto, scrivi il nome esatto del file."
read -rp "Nome file icona PNG: " ICON_FILE

[ -n "$ICON_FILE" ] || error_exit "Nome file icona vuoto"
ICON_PATH="$EXPECTED_DIR/$ICON_FILE"
check_file_exists "$ICON_PATH" "Icona PNG"
echo "OK: trovata icona $ICON_PATH"
pause_step

echo "STEP 3 - Inserisci nella cartella $EXPECTED_DIR il file desktop di DataFlow."
echo "Quando l'hai fatto, scrivi il nome esatto del file."
read -rp "Nome file .desktop: " DESKTOP_FILE

[ -n "$DESKTOP_FILE" ] || error_exit "Nome file desktop vuoto"
DESKTOP_PATH="$EXPECTED_DIR/$DESKTOP_FILE"
check_file_exists "$DESKTOP_PATH" "File desktop"
echo "OK: trovato file desktop $DESKTOP_PATH"
pause_step

echo "STEP 4 - Preparazione cartella AppDir dentro $EXPECTED_DIR ..."
rm -rf "$APPDIR" || error_exit "Impossibile rimuovere AppDir precedente"
mkdir -p "$APPDIR/usr/bin" || error_exit "Impossibile creare $APPDIR/usr/bin"
mkdir -p "$APPDIR/usr/share" || error_exit "Impossibile creare $APPDIR/usr/share"
echo "OK: AppDir pronta."
pause_step

echo "STEP 5 - Creazione file AppRun..."
cat > "$APPDIR/AppRun" <<'EOF'
#!/bin/sh
HERE="$(dirname "$(readlink -f "${0}")")"
exec "${HERE}/usr/share/dataflow/venv/bin/python" "${HERE}/usr/share/dataflow/DataFlow 2.0.0.py" "$@"
EOF

chmod +x "$APPDIR/AppRun" || error_exit "Impossibile rendere eseguibile AppRun"
echo "OK: AppRun creato."
pause_step

echo "STEP 6 - Copia icona e file desktop nella root di AppDir..."
cp "$ICON_PATH" "$APPDIR/dataflow.png" || error_exit "Impossibile copiare l'icona"
cp "$DESKTOP_PATH" "$APPDIR/dataflow.desktop" || error_exit "Impossibile copiare il file desktop"
echo "OK: icona e desktop copiati."
pause_step

echo "STEP 7 - Estrazione del .deb dentro AppDir..."
dpkg-deb -x "$DEB_PATH" "$APPDIR" || error_exit "Estrazione del .deb fallita"
echo "OK: pacchetto estratto."
pause_step

echo "STEP 8 - Verifica contenuti principali..."
check_file_exists "$APPDIR/usr/bin/dataflow" "Launcher usr/bin/dataflow"
check_file_exists "$APPDIR/usr/share/dataflow" "Cartella applicazione usr/share/dataflow"
check_file_exists "$APPDIR/usr/share/dataflow/venv/bin/python" "Python del venv"
check_file_exists "$APPDIR/usr/share/dataflow/DataFlow 2.0.0.py" "Script principale"
echo "OK: file principali trovati."
pause_step

echo "STEP 9 - Elenco contenuti principali trovati..."
echo
echo "Contenuto di $APPDIR/usr/share/dataflow:"
ls "$APPDIR/usr/share/dataflow" || error_exit "Impossibile leggere la cartella applicazione"
echo
echo "Contenuto di $APPDIR/usr/share/dataflow/venv/bin:"
ls "$APPDIR/usr/share/dataflow/venv/bin" || error_exit "Impossibile leggere il venv"
pause_step

echo "STEP 10 - Creazione AppImage finale dentro $EXPECTED_DIR ..."
rm -f "$OUTPUT_APPIMAGE"
ARCH=x86_64 "$APPIMAGE_TOOL" "$APPDIR" || error_exit "Creazione AppImage fallita"
echo "OK: AppImage creata."
pause_step

echo "STEP 11 - Verifica file finale..."
check_file_exists "$OUTPUT_APPIMAGE" "AppImage finale"
chmod +x "$OUTPUT_APPIMAGE" || error_exit "Impossibile rendere eseguibile l'AppImage"
echo "OK: AppImage pronta in $OUTPUT_APPIMAGE"
echo
ls -lh "$OUTPUT_APPIMAGE"
pause_step

echo "STEP 12 - Vuoi avviare subito l'AppImage per testarla?"
read -rp "Digita s per sì, qualsiasi altro tasto per uscire: " RUN_NOW

if [ "$RUN_NOW" = "s" ] || [ "$RUN_NOW" = "S" ]; then
    echo "Avvio AppImage..."
    "$OUTPUT_APPIMAGE" || error_exit "L'AppImage si è chiusa con errore"
else
    echo "Test finale saltato."
fi

echo
echo "========================================"
echo " Operazione completata con successo "
echo "========================================"
echo "File creato: $OUTPUT_APPIMAGE"
