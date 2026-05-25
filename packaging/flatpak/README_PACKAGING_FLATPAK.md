# DataFlow Flatpak packaging draft

## Scopo

Questa cartella contiene una prima bozza di packaging Flatpak per DataFlow.
Non modifica il layout del progetto e non richiede modifiche al codice
applicativo Python.

App-id provvisorio:

```text
io.github.sorguido.DataFlow
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

- `io.github.sorguido.DataFlow.yml`: manifest Flatpak iniziale.
- `io.github.sorguido.DataFlow.desktop`: launcher desktop Linux.
- `io.github.sorguido.DataFlow.metainfo.xml`: metainfo/AppStream provvisorio.
- `run-dataflow.sh`: wrapper installato come `/app/bin/dataflow`.
- `README_PACKAGING_FLATPAK.md`: note operative e limiti della bozza.

## Prerequisiti locali su Fedora

```bash
sudo dnf install flatpak flatpak-builder
flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
flatpak install flathub org.freedesktop.Platform//24.08 org.freedesktop.Sdk//24.08
```

Nota: questa bozza usa `org.freedesktop.Platform` e `org.freedesktop.Sdk`
24.08 come base iniziale. La scelta va validata per Tkinter/Tcl/Tk.

## Build ipotetica

Da root repository:

```bash
flatpak-builder --force-clean --user --install-deps-from=flathub build-flatpak packaging/flatpak/io.github.sorguido.DataFlow.yml
```

Per installare localmente dopo la build:

```bash
flatpak-builder --force-clean --user --install --install-deps-from=flathub build-flatpak packaging/flatpak/io.github.sorguido.DataFlow.yml
```

La sezione `python-dependencies` del manifest usa temporaneamente `pip`.
Questo e' comodo per una prova locale, ma non e' una soluzione pronta per
Flathub. Uno step successivo dovrebbe sostituirla con moduli generati in modo
riproducibile, ad esempio tramite `flatpak-pip-generator` o equivalente.

## Test ipotetici

Avvio dal sandbox:

```bash
flatpak run io.github.sorguido.DataFlow
```

Avvio con log dettagliato Flatpak:

```bash
flatpak run --verbose io.github.sorguido.DataFlow
```

Shell di debug nel sandbox:

```bash
flatpak run --command=sh io.github.sorguido.DataFlow
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
fornire Python con `_tkinter` funzionante e le librerie Tcl/Tk necessarie.
Questo punto non e' risolto dal manifest in modo definitivo.

Se `python3 -c "import tkinter"` fallisce nel sandbox, le opzioni realistiche
sono:

- cambiare runtime/base piu adatto;
- includere Python/Tcl/Tk come dipendenze di build/runtime;
- creare un runtime packaging piu specifico.

Queste opzioni sono volutamente lasciate fuori da questa prima bozza.

### Dipendenze Python

Da `requirements.txt` sono considerate runtime:

- `openpyxl==3.1.5`
- `Pillow==12.1.1`
- `reportlab==4.2.2`
- `tkcalendar==1.6.1`
- `tksheet==7.6.0`
- `tkinterdnd2==0.4.2`

`polib==1.2.0` non e' incluso nella bozza runtime perche risulta usato dal
tool di sviluppo `development/dev_tools/compile_translations.py`, mentre
l'app usa gia i file `.mo` presenti in `locale/`.

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

## Cosa non e' ancora risolto

- Verifica reale di Tkinter/Tcl/Tk nel runtime `org.freedesktop.Platform`.
- Packaging riproducibile delle dipendenze Python senza accesso network in build.
- Strategia definitiva per path dati sotto Flatpak.
- Permessi filesystem minimi per import/export/backup/allegati.
- Integrazione portali per apertura URL/file.
- Validazione AppStream completa per pubblicazione.
- Set icone Linux completo in layout `hicolor`.
- Eventuale test su Wayland rispetto a X11.

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

