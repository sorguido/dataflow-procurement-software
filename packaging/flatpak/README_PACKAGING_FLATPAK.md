# DataFlow Flatpak packaging

## Purpose

This directory contains the Flatpak packaging for DataFlow.
The packaging does not change the project layout and does not require changes
to the Python application code.

Guiding principle: the package adapts to DataFlow, not DataFlow to the package.

App-id:

```text
io.github.sorguido.dataflow-procurement-software
```

Expected command inside the sandbox:

```text
dataflow
```

Application entrypoint:

```text
dataflow.py
```

## Included files

- `io.github.sorguido.dataflow-procurement-software.yml`: Flatpak manifest.
- `io.github.sorguido.dataflow-procurement-software.desktop`: Linux desktop launcher.
- `io.github.sorguido.dataflow-procurement-software.metainfo.xml`: metainfo/AppStream.
- `run-dataflow.sh`: wrapper installed as `/app/bin/dataflow`.
- `README_PACKAGING_FLATPAK.md`: operational notes and draft limitations.

## Local prerequisites on Fedora

```bash
sudo dnf install flatpak flatpak-builder
flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
flatpak install flathub org.freedesktop.Platform//24.08 org.freedesktop.Sdk//24.08
```

Note: this draft uses `org.freedesktop.Platform` and `org.freedesktop.Sdk`
24.08 as its initial base. Tcl/Tk 8.6.14 are built inside `/app`
to avoid the minimal Tk provided by the Freedesktop runtime.

## Local build

From the repository root, to build and install locally:

```bash
rm -rf /tmp/dataflow-flatpak-build /tmp/dataflow-flatpak-state

flatpak-builder \
  --force-clean \
  --user \
  --install \
  --state-dir=/tmp/dataflow-flatpak-state \
  /tmp/dataflow-flatpak-build \
  packaging/flatpak/io.github.sorguido.dataflow-procurement-software.yml
```

To generate a manual distributable Flatpak bundle:

```bash
rm -rf /tmp/dataflow-flatpak-build \
       /tmp/dataflow-flatpak-state \
       /tmp/dataflow-flatpak-repo \
       /tmp/dataflow-flatpak-dist

mkdir -p /tmp/dataflow-flatpak-dist

flatpak-builder \
  --force-clean \
  --user \
  --repo=/tmp/dataflow-flatpak-repo \
  --state-dir=/tmp/dataflow-flatpak-state \
  /tmp/dataflow-flatpak-build \
  packaging/flatpak/io.github.sorguido.dataflow-procurement-software.yml

flatpak build-bundle \
  --runtime-repo=https://flathub.org/repo/flathub.flatpakrepo \
  /tmp/dataflow-flatpak-repo \
  /tmp/dataflow-flatpak-dist/DataFlow-2.3.0-x86_64.flatpak \
  io.github.sorguido.dataflow-procurement-software
```

The bundle does not include the Freedesktop runtime: `--runtime-repo` tells
Flatpak where to retrieve it, normally from Flathub.

The Python runtime dependencies are no longer installed with a direct
`pip install` in the main manifest. They are referenced through
`python3-requirements.json`, generated with `flatpak-pip-generator`, which
contains the URLs and SHA-256 values for the required sources/wheels. This
makes the manifest more reproducible and removes the need for the
`--share=network` build argument for the Python module.

## Tests and smoke tests

Start from the sandbox:

```bash
flatpak run io.github.sorguido.dataflow-procurement-software
```

Start with detailed Flatpak logging:

```bash
flatpak run --verbose io.github.sorguido.dataflow-procurement-software
```

Debug shell inside the sandbox:

```bash
flatpak run --command=sh io.github.sorguido.dataflow-procurement-software
```

Minimum manual checks:

- first launch
- Tkinter window loading
- license, language, and user identity dialogs
- SQLite database creation
- window icon loading
- translation loading from `locale/`
- Excel export
- PDF export
- attachment handling
- external link opening
- file/template opening with an external application
- attachment drag and drop, if available

## Known issues

### Tkinter, Tcl, and Tk

DataFlow uses Tkinter as its main GUI. The selected Flatpak runtime must
provide Python with a working `_tkinter` module. In this draft, Tcl/Tk are not
taken from the runtime: the manifest builds Tcl 8.6.14 and Tk 8.6.14 inside
`/app`.

The reason is diagnostic and practical: the Freedesktop runtime
`libtk8.6.so` may be built without Xft/fontconfig/freetype. In that case
Tkinter starts, but Tk sees only a few core X11 fonts and `TkDefaultFont`
often becomes `fixed`. By building Tk in the bundle with `--enable-xft`, Tk
should link to:

- `libXft`
- `libfontconfig`
- `libfreetype`
- `libXrender`

The `run-dataflow.sh` wrapper prepends `/app/lib` to `LD_LIBRARY_PATH` and sets
`TCL_LIBRARY`/`TK_LIBRARY` only when the installed directories exist. This makes
Python/_tkinter prefer the Tcl/Tk libraries installed in `/app` over the ones
provided by the runtime.

To verify that the bundled Tk is the expected one:

```bash
flatpak run --command=sh io.github.sorguido.dataflow-procurement-software -c 'ldd /app/lib/libtk8.6.so | grep -E "libXft|libfontconfig|libfreetype|libXrender"'
```

To verify which default font Tk sees and how many font families are available:

```bash
flatpak run --command=python3 io.github.sorguido.dataflow-procurement-software -c 'import tkinter as tk; import tkinter.font as f; root=tk.Tk(); print("patchlevel:", root.tk.call("info", "patchlevel")); print("TkDefaultFont:", f.nametofont("TkDefaultFont").actual()); families=sorted(f.families()); print("font_count:", len(families)); print("sample:", families[:40]); root.destroy()'
```

If `python3 -c "import tkinter"` fails inside the sandbox even with Tcl/Tk in
`/app`, the realistic options are:

- switch to a more suitable runtime/base;
- also include Python/_tkinter in a controlled way;
- create a more specific packaging runtime.

These options are intentionally left out of this draft to avoid introducing
invasive changes.

### Python dependencies

The following dependencies from `requirements.txt` are considered runtime
dependencies and are included in `python3-requirements.json`:

- `openpyxl==3.1.5`
- `Pillow==12.1.1`
- `reportlab==4.2.2`
- `tkcalendar==1.6.1`
- `tksheet==7.6.0`
- `tkinterdnd2==0.4.2`

`polib==1.2.0` is not included in the runtime draft because it appears to be
used by the development tool `development/dev_tools/compile_translations.py`,
while the application already uses the `.mo` files present in `locale/`.

The main manifest keeps this order:

```text
tcl
tk
python3-requirements.json
dataflow
```

This is intentional: Tcl/Tk bundled in `/app` must be built before the Python
dependencies and before the application files are installed.

### File system and sandbox

DataFlow must preserve the same behavior as the native Linux and Windows
versions: exports, backups, attachments, and user-selected directories must be
able to write to the real host user home directory.

For this reason, the manifest grants:

```text
--filesystem=home
```

This choice is intentional. A more restrictive configuration based only on
portals or individual directories would require application-level path handling
changes or would change DataFlow's expected desktop-native behavior.

Verify installed permissions:

```bash
flatpak info --show-permissions io.github.sorguido.dataflow-procurement-software
```

Expected output in the `[Context]` block:

```text
shared=ipc;
sockets=x11;
filesystems=home;
```

#### XDG permission test without full-home access

A temporary test was performed by removing `--filesystem=home` and using
stricter permissions:

- `--filesystem=xdg-documents:create`
- `--filesystem=xdg-download:create`
- `--filesystem=xdg-desktop:create`

Lint result:

- the `finish-args-home-filesystem-access` error disappears;
- only the already documented runtime 25.08 warning remains.

Functional result:

- the sandbox correctly sees and writes to the XDG Documents, Downloads, and
  Desktop directories;
- DataFlow starts, but does not preserve the expected desktop-native behavior;
- a new alternative/sandboxed database is created;
- a new RfQ created during the test is not persistent after restarting the
  application.

Conclusion:

for DataFlow 2.3.0, XDG-only permissions are not sufficient. Without dedicated
application-level path handling changes for paths, databases, runtime
configuration, and user-selected files, `--filesystem=home` remains necessary
to preserve the behavior of the native desktop version.

For this reason, the `finish-args-home-filesystem-access` lint error is
documented and accepted for the manual DataFlow 2.3.0 bundle. Any future
Flathub-compatible variant that is more sandbox-friendly requires separate
analysis and explicit application changes.

### Browser and file opening

The code uses `webbrowser.open(...)` and, in one case, `xdg-open`. In Flatpak,
the correct solution should go through desktop portals. This draft does not
introduce invasive wrappers and only documents the risk.

### Drag and drop

`tkinterdnd2` may depend on Tcl/Tk components and desktop integrations that are
not always available inside the sandbox. The app already has a manual fallback
through the file selection button.

### Tcl/Tk bundle

The Tcl/Tk 8.6.14 sources are downloaded from SourceForge with explicit
SHA-256 values in the manifest. The local build uses:

```text
Tcl: tcl8.6.14-src.tar.gz
Tk:  tk8.6.14-src.tar.gz
```

Tk is configured with:

```text
--enable-xft
```

Known limitations:

- the Tk SHA-256 is taken from the upstream tarball republished as an
  Ubuntu/Launchpad source; before a Flathub submission it must be rechecked
  against the actual download used by `flatpak-builder`;
- the build checks `ldd /app/lib/libtk8.6.so` during the `tk` module, but it
  must still be tested in the installed runtime;
- `_tkinter` remains the one provided by the runtime Python; this draft only
  changes the preferred Tcl/Tk libraries at runtime;
- Wayland has not been enabled at this stage.

## Flatpak lint status

The manifest is functionally validated for the current target: a manual
distributable Flatpak bundle.

flatpak-builder-lint intentionally reports:

    finish-args-home-filesystem-access

Reason: the manifest grants:

    --filesystem=home

This choice is intentional. DataFlow must preserve the behavior of the native
desktop version: Excel/PDF export, database backups, attachments, and
user-selected directories must write to the real host user home directory.

For a possible Flathub target, this point will need to be reassessed. Future
options are:

- desktop portal integration;
- stricter permissions on specific XDG directories;
- exception request or justification, if applicable.

The following warning is not addressed at this stage:

    runtime-update-available-to-org.freedesktop.Platform-25.08

Runtime 24.08 remains the runtime validated for Tcl/Tk, fonts, Python
dependencies, build, and application startup.

Runtime 25.08 note:

A temporary test with org.freedesktop.Platform//25.08 and
org.freedesktop.Sdk//25.08 completed the build successfully, but the runtime
was not adopted for DataFlow 2.3.0.

Technical reason:

- 24.08 uses Python 3.12 and includes tkinter / _tkinter;
- 25.08 uses Python 3.13, but in the tested runtime it does not expose
  tkinter / _tkinter;
- DataFlow is a Tkinter desktop app;
- tkcalendar, tksheet, and tkinterdnd2 also depend on Tkinter.

The Tcl/Tk 8.6.14 bundle built in /app solves the Tcl/Tk library and font
issues, but it cannot replace the Python _tkinter module, which remains
provided by the runtime Python.

For this reason, the runtime-update-available-to-org.freedesktop.Platform-25.08
warning is documented and accepted for the DataFlow 2.3.0 bundle based on
24.08.



## What is not resolved yet

- Periodic verification of the `python3-requirements.json` module when the
  runtime Python version or dependencies change.
- Future evaluation of a newer Freedesktop runtime compared with `24.08`.
- Portal integration for opening URLs/files, if required in a later phase.
- Full AppStream validation for publication.
- Complete Linux icon set in the `hicolor` layout.
- Possible Wayland testing compared with X11.
- Final confirmation of source checksums against the actual mirrors used by
  `flatpak-builder`.

## Rollback

This draft is confined to `packaging/flatpak/`.

To remove it before a commit:

```bash
rm -rf packaging/flatpak
```

To verify that no other files were touched:

```bash
git status --short
git diff --stat
```
