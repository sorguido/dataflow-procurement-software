#!/usr/bin/env bash
set -euo pipefail

APP_NAME="dataflow"
APP_DISPLAY_NAME="DataFlow Procurement Software"
VERSION="2.0.1"
ARCH="amd64"
MAIN_SCRIPT="dataflow.py"

BUILD_DIR="dataflow-deb"
INSTALL_ROOT="$BUILD_DIR/usr"
APP_DIR="$INSTALL_ROOT/share/dataflow"
BIN_DIR="$INSTALL_ROOT/bin"
DESKTOP_DIR="$INSTALL_ROOT/share/applications"
ICON_DIR="$INSTALL_ROOT/share/icons/hicolor/256x256/apps"

echo "==> Pulizia build precedente..."
rm -rf "$BUILD_DIR"
rm -f ./*.deb

echo "==> Creazione struttura cartelle..."
mkdir -p "$APP_DIR" "$BIN_DIR" "$DESKTOP_DIR" "$ICON_DIR"

echo "==> Verifica entrypoint..."
if [ ! -f "$MAIN_SCRIPT" ]; then
    echo "ERRORE: file principale non trovato: $MAIN_SCRIPT"
    exit 1
fi

if [ ! -f "requirements.txt" ]; then
    echo "ERRORE: requirements.txt non trovato."
    exit 1
fi

if [ ! -f "add_data/Logo_150x150.png" ]; then
    echo "ERRORE: icona non trovata: add_data/Logo_150x150.png"
    exit 1
fi

echo "==> Verifica presenza rsync..."
if ! command -v rsync >/dev/null 2>&1; then
    echo ""
    echo "ERRORE: rsync non trovato."
    echo ""
    echo "Installa con:"
    echo "sudo apt install rsync"
    exit 1
fi

echo "==> Copia progetto..."
rsync -av ./ "$APP_DIR/" \
  --exclude '.git/' \
  --exclude '.github/' \
  --exclude '.venv/' \
  --exclude 'venv/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  --exclude '*.pyo' \
  --exclude '*.deb' \
  --exclude "$BUILD_DIR/" \
  --exclude 'build/' \
  --exclude '.mypy_cache/' \
  --exclude '.pytest_cache/' \
  --exclude '.idea/' \
  --exclude '.vscode/' \
  --exclude '.DS_Store' \
  --exclude '*.backup_broken'

echo "==> Creazione virtual environment..."
python3 -m venv "$APP_DIR/venv"

echo "==> Aggiornamento pip nella venv..."
"$APP_DIR/venv/bin/pip" install --upgrade pip setuptools wheel

echo "==> Installazione dipendenze Python..."
"$APP_DIR/venv/bin/pip" install -r "$APP_DIR/requirements.txt"

echo "==> Creazione launcher..."
cat > "$BIN_DIR/dataflow" << EOF
#!/usr/bin/env bash
exec /usr/share/dataflow/venv/bin/python "/usr/share/dataflow/$MAIN_SCRIPT"
EOF
chmod +x "$BIN_DIR/dataflow"

echo "==> Creazione desktop entry..."
cat > "$DESKTOP_DIR/dataflow.desktop" << 'EOF'
[Desktop Entry]
Version=1.0
Name=DataFlow Procurement Software
GenericName=Procurement Software
Comment=Procurement management tool
Exec=/usr/bin/dataflow
Icon=dataflow
Terminal=false
Type=Application
Categories=Office;
Keywords=procurement;rfq;purchase;acquisti;
StartupNotify=true
EOF
chmod 644 "$DESKTOP_DIR/dataflow.desktop"

echo "==> Installazione icona..."
cp "add_data/Logo_150x150.png" "$ICON_DIR/dataflow.png"
chmod 644 "$ICON_DIR/dataflow.png"

echo "==> Verifica presenza fpm..."
if ! command -v fpm >/dev/null 2>&1; then
    echo ""
    echo "ERRORE: fpm non trovato."
    echo ""
    echo "Installa con:"
    echo "sudo apt install ruby ruby-dev build-essential"
    echo "sudo gem install --no-document fpm"
    exit 1
fi

echo "==> Generazione pacchetto .deb..."
fpm -s dir -t deb \
  -n "$APP_NAME" \
  -v "$VERSION" \
  -a "$ARCH" \
  --description "$APP_DISPLAY_NAME" \
  --maintainer "Guido Sorarù" \
  --license "GPL-3.0" \
  --url "https://github.com/sorguido/dataflow-procurement-software" \
  --prefix=/ \
  --depends python3 \
  --depends python3-tk \
  -C "$BUILD_DIR" \
  .

echo ""
echo "==> Pacchetto creato:"
ls -lh ./*.deb

echo ""
echo "==> Installazione suggerita:"
echo "sudo apt install ./dataflow_${VERSION}_${ARCH}.deb"
