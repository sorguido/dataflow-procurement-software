# DataFlow Procurement Software

DataFlow Procurement Software is a desktop application for managing Requests for Quotation (RfQ) in purchasing workflows.

Originally developed for Windows and also published on the Microsoft Store, the project is now being released as an open-source Linux edition under the **GNU GPLv3** license.

The application is written in **Python** with a **Tkinter** GUI and uses **SQLite** as its local database engine. The current Linux port already includes cross-platform path handling, Linux window icon support, multilingual support (Italian and English), Excel import/export, attachment management, purchase order tracking, notes, and SQDC analysis support. The codebase also includes a dedicated database manager with SQLite/WAL support and logic for aggregating data across multiple user databases.

---

## Highlights

- Desktop application for procurement and RfQ management
- Python + Tkinter graphical interface
- SQLite database backend
- Excel import/export with `openpyxl`
- Attachment handling
- Purchase order (PO) tracking
- Notes management
- SQDC analysis export/save workflow
- English and Italian language support via gettext/polib
- Linux-compatible port with fixes for platform-specific behavior
- Existing Windows distribution on Microsoft Store

---

## Screenshots

![Main Window](docs/screenshots/main_window.png)

---

## Project status

The Linux version is currently the open-source edition of the project.

Recent work on the port includes:

- removal/fix of Windows-specific paths
- Linux-compatible window icon handling
- Tkinter fix related to `grab_set()` placement after `wait_visibility()`
- update of licensing from Freeware to GNU GPLv3 for the Linux release
- multilingual updates in both the application and documentation
- footer/documentation cleanup aligned with the new license

---

## Tech stack

- **Language:** Python
- **GUI:** Tkinter
- **Database:** SQLite
- **IDE used during development:** Visual Studio Code
- **Main file:** `DataFlow 2.0.0.py`

### Main dependencies

- `openpyxl`
- `Pillow`
- `polib`
- `tkcalendar`
- `tksheet`

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/sorguido/dataflow-procurement-software.git
cd dataflow-procurement-software
```

### 2. Create a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the application

```bash
python3 "DataFlow 2.0.0.py"
```

---

## requirements.txt example

```txt
openpyxl
Pillow
polib
tkcalendar
tksheet
```

You can pin versions later after testing the Linux release more broadly.

---

## License

This project is released under the **GNU General Public License v3.0**.

The complete license text must be included in the repository in the `LICENSE` file.

---

## Windows version note

A Windows version of DataFlow also exists and has been published on the Microsoft Store  
```https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=it-IT&gl=IT)```

At the moment, the Linux edition is the open-source GNU GPLv3 release. If future Windows releases are aligned with the same licensing model, they may also be distributed through this repository or a related packaging workflow.

---

## Contributing

DataFlow is now available as an open-source project.

The Linux version of the application has been released under the GNU GPLv3 license and the source code is available on GitHub.

Developers interested in improving or adapting the software, including future Windows versions, are welcome to contribute.

---

## Repository link

Once the GitHub repository is created, replace the placeholders in this README with the final public URL:

```text
https://github.com/sorguido/dataflow-procurement-software
```
