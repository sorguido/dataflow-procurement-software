# 08 – Settings and Maintenance

## Opening Settings

Click **⚙️ Settings** in the toolbar. The **"Settings and Maintenance"** window opens.

---

## DataFlow Location

### Viewing the Current Location

The **"DataFlow Standard Location"** section shows the path of the active folder. This folder contains the database and attachments.

### Changing the Location

1. Click **📁 Change DataFlow Location...**
2. Select the new folder in the dialog.
3. Click **OK**.
4. A message informs you that the application must be restarted for the change to take effect.

> **Warning:** changing the location does **not** automatically copy the existing database. Before restarting, manually copy the entire `DataFlow/` folder to the new location.

---

## Manual Backup

To create a backup copy of the database at any time:

1. Click **💾 Manual Backup...**
2. Choose the destination folder and file name in the save dialog.
3. The `dataflow_db.db` file is copied to the chosen location.

It is recommended to perform a manual backup before any critical operation (e.g. migrating to a new server, changing the database path).

---

## Daily Automatic Backup

DataFlow can automatically back up the database every day at a configured time.

### Configuring Automatic Backup

1. Enable the **"Daily Automatic Backup"** checkbox.
2. Set the **Hour** (00–23) in the dropdown.
3. In the **"Save to:"** field, click **📁 Choose...** to select the backup destination folder.
4. Click **💾 Save Backup Settings**.

DataFlow retains a maximum of **3 automatic backups**. When a new one is created, the oldest is deleted automatically.

Settings are saved in `config.ini` under the `[AutoBackup]` section.

---

## Interface Language

1. In the **"Language"** dropdown, select:
   - **Italiano**
   - **English**
2. Click **💾 Save Language**.
3. A message informs you that a restart is required for the language change to take effect.

The language is saved in `config.ini` under the `language` key. After restarting, all menus, labels, and windows will be displayed in the selected language.

---

## Configuration File

Application settings are stored in the `config.ini` file, located in the same folder as the database (`DataFlow/`). The main sections are:

| Section | Content |
|---------|---------|
| `[Settings]` | Language, database path, payment terms financial coefficient |
| `[AutoBackup]` | Enable/disable, hour, backup destination folder |
| `[User]` | First name, last name, generated username |

In case of application issues, a technician can read these values to diagnose the configuration.

---

## Log File Paths

Diagnostic logs are written automatically to:

- **Windows:** `%LOCALAPPDATA%\DataFlow\dataflow.log`
- **Linux:** `~/.local/share/DataFlow/dataflow.log`

See the [Logs and Diagnostics](13-Logs-and-Diagnostics.md) section for more details.
