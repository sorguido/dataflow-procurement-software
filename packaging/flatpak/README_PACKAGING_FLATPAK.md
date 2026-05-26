# DataFlow Flatpak packaging draft

## Scopo

Questa cartella contiene una prima bozza di packaging Flatpak per DataFlow.
Non modifica il layout del progetto e non richiede modifiche al codice
applicativo Python.

App-id provvisorio:

```text
io.github.sorguido.dataflow-procurement-software
```

Comando previsto nel sandbox:

```text
dataflow
```

Entrypoint applicativo:

```text
dataflow.py
```

## File inclusi

- `io.github.sorguido.dataflow-procurement-software.yml`: manifest Flatpak iniziale.
- `io.github.sorguido.dataflow-procurement-software.desktop`: launcher desktop Linux.
- `io.github.sorguido.dataflow-procurement-software.metainfo.xml`: metainfo/AppStream provvisorio.
- `run-dataflow.sh`: wrapper installato come `/app/bin/dataflow`.
- `README_PACKAGING_FLATPAK.md`: note operative e limiti della bozza.

## Prerequisiti locali su Fedora

```bash
sudo dnf install flatpak flatpak-builder
flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
flatpak install flathub org.freedesktop.Platform//24.08 org.freedesktop.Sdk//24.08
```

Nota: questa bozza usa `org.freedesktop.Platform` e `org.freedesktop.Sdk`
24.08 come base iniziale. Tcl/Tk 8.6.14 vengono compilati dentro `/app`
per evitare il Tk minimale del runtime Freedesktop.

## Build ipotetica

Da root repository:

```bash
flatpak-builder --force-clean --user --install-deps-from=flathub build-flatpak packaging/flatpak/io.github.sorguido.dataflow-procurement-software.yml
```

Per installare localmente dopo la build:

```bash
flatpak-builder --force-clean --user --install --install-deps-from=flathub build-flatpak packaging/flatpak/io.github.sorguido.dataflow-procurement-software.yml
```

Le dipendenze Python runtime non vengono piu installate con un `pip install`
diretto nel manifest principale. Sono referenziate tramite
`python3-requirements.json`, generato con `flatpak-pip-generator`, che contiene
URL e SHA-256 dei sorgenti/wheel necessari. Questo mantiene il manifest piu
riproducibile e rimuove la necessita del build arg `--share=network` per il
modulo Python.

## Test ipotetici

Avvio dal sandbox:

```bash
flatpak run io.github.sorguido.dataflow-procurement-software
```

Avvio con log dettagliato Flatpak:

```bash
flatpak run --verbose io.github.sorguido.dataflow-procurement-software
```

Shell di debug nel sandbox:

```bash
flatpak run --command=sh io.github.sorguido.dataflow-procurement-software
```

Controlli manuali minimi:

- primo avvio
- caricamento finestra Tkinter
- dialog licenza, lingua e identita utente
- creazione database SQLite
- caricamento icona finestra
- caricamento traduzioni da `locale/`
- export Excel
- export PDF
- gestione allegati
- apertura link esterni
- apertura file/template con applicazione esterna
- drag and drop allegati, se disponibile

## Problemi noti

### Tkinter, Tcl e Tk

DataFlow usa Tkinter come GUI principale. Il runtime Flatpak scelto deve
fornire Python con `_tkinter` funzionante. In questa bozza Tcl/Tk non vengono
presi dal runtime: il manifest compila Tcl 8.6.14 e Tk 8.6.14 dentro `/app`.

Il motivo e' diagnostico e pratico: il `libtk8.6.so` del runtime Freedesktop
puo essere compilato senza Xft/fontconfig/freetype. In quel caso Tkinter parte,
ma Tk vede pochi font core X11 e `TkDefaultFont` diventa spesso `fixed`.
Compilando Tk nel bundle con `--enable-xft`, Tk dovrebbe collegarsi a:

- `libXft`
- `libfontconfig`
- `libfreetype`
- `libXrender`

Il wrapper `run-dataflow.sh` prepende `/app/lib` a `LD_LIBRARY_PATH` e imposta
`TCL_LIBRARY`/`TK_LIBRARY` solo se le directory installate esistono. Questo
serve a far preferire a Python/_tkinter le librerie Tcl/Tk installate in `/app`
rispetto a quelle del runtime.

Per verificare che il Tk bundle sia quello corretto:

```bash
flatpak run --command=sh io.github.sorguido.dataflow-procurement-software -c 'ldd /app/lib/libtk8.6.so | grep -E "libXft|libfontconfig|libfreetype|libXrender"'
```

Per verificare quale font predefinito vede Tk e quante famiglie font sono
disponibili:

```bash
flatpak run --command=python3 io.github.sorguido.dataflow-procurement-software -c 'import tkinter as tk; import tkinter.font as f; root=tk.Tk(); print("patchlevel:", root.tk.call("info", "patchlevel")); print("TkDefaultFont:", f.nametofont("TkDefaultFont").actual()); families=sorted(f.families()); print("font_count:", len(families)); print("sample:", families[:40]); root.destroy()'
```

Se `python3 -c "import tkinter"` fallisce nel sandbox anche con Tcl/Tk in
`/app`, le opzioni realistiche sono:

- cambiare runtime/base piu adatto;
- includere anche Python/_tkinter in modo controllato;
- creare un runtime packaging piu specifico.

Queste opzioni sono volutamente lasciate fuori da questa bozza per non
introdurre cambi invasivi.

### Dipendenze Python

Da `requirements.txt` sono considerate runtime e incluse in
`python3-requirements.json`:

- `openpyxl==3.1.5`
- `Pillow==12.1.1`
- `reportlab==4.2.2`
- `tkcalendar==1.6.1`
- `tksheet==7.6.0`
- `tkinterdnd2==0.4.2`

`polib==1.2.0` non e' incluso nella bozza runtime perche risulta usato dal
tool di sviluppo `development/dev_tools/compile_translations.py`, mentre
l'app usa gia i file `.mo` presenti in `locale/`.

Il manifest principale mantiene l'ordine:

```text
tcl
tk
python3-requirements.json
dataflow
```

Questo e' intenzionale: Tcl/Tk bundled in `/app` devono essere costruiti prima
delle dipendenze Python e prima dell'installazione dei file applicativi.

### File system e sandbox

Il manifest mantiene permessi runtime prudenti. Non concede accesso ampio a
`home` in questa prima bozza.

DataFlow oggi usa cartelle utente come:

```text
~/.local/share/DataFlow/
~/DataFlow_<username>/
```

Nel sandbox Flatpak questo comportamento va testato con attenzione. Import,
export, backup, allegati e scelta cartelle potrebbero richiedere portali o
permessi filesystem aggiuntivi. Eventuali permessi temporanei sono commentati
nel manifest o descritti qui, ma non vanno considerati una scelta definitiva.

### Apertura browser e file

Il codice usa `webbrowser.open(...)` e in un caso `xdg-open`. In Flatpak la
soluzione corretta dovrebbe passare dai portali desktop. Questa bozza non
introduce wrapper invasivi e si limita a documentare il rischio.

### Drag and drop

`tkinterdnd2` puo dipendere da componenti Tcl/Tk e integrazioni desktop non
sempre disponibili nel sandbox. L'app ha gia un fallback manuale tramite
pulsante di selezione file.

### Tcl/Tk bundle

I sorgenti Tcl/Tk 8.6.14 sono scaricati da SourceForge con SHA-256 espliciti
nel manifest. La build locale usa:

```text
Tcl: tcl8.6.14-src.tar.gz
Tk:  tk8.6.14-src.tar.gz
```

Tk viene configurato con:

```text
--enable-xft
```

Limiti noti:

- lo SHA-256 di Tk e' tratto dal tarball upstream ripubblicato come sorgente
  Ubuntu/Launchpad; prima di una submission Flathub va ricontrollato contro
  il download effettivo usato da `flatpak-builder`;
- la build verifica `ldd /app/lib/libtk8.6.so` durante il modulo `tk`, ma va
  comunque testata nel runtime installato;
- `_tkinter` resta quello fornito dal runtime Python; questa bozza cambia solo
  le librerie Tcl/Tk preferite a runtime;
- Wayland non e' stato abilitato in questa fase.

## Cosa non e' ancora risolto

- Verifica reale di `_tkinter` con Tcl/Tk 8.6.14 installati in `/app`.
- Verifica periodica del modulo `python3-requirements.json` quando cambiano
  versioni Python runtime o dipendenze.
- Strategia definitiva per path dati sotto Flatpak.
- Permessi filesystem minimi per import/export/backup/allegati.
- Integrazione portali per apertura URL/file.
- Validazione AppStream completa per pubblicazione.
- Set icone Linux completo in layout `hicolor`.
- Eventuale test su Wayland rispetto a X11.
- Conferma finale dei checksum sorgenti rispetto ai mirror effettivi usati da
  `flatpak-builder`.

## Rollback

Questa bozza e' confinata in `packaging/flatpak/`.

Per rimuoverla prima di un commit:

```bash
rm -rf packaging/flatpak
```

Per verificare che non siano stati toccati altri file:

```bash
git status --short
git diff --stat
```
