# 01 – Getting Started

## System Requirements

DataFlow runs on **Windows 10/11** and **Linux** (any distribution with a graphical desktop). No internet connection is required. For shared multi-user use, all computers simply need access to the same folder on a server or NAS.

---

## First Launch

On the very first launch, DataFlow displays an identity setup window. You must enter your **First Name** and **Last Name** before the application can be used.

### Setting Up Your Identity

1. When the **"User Identity"** window appears, enter your first name in the **First Name** field and your last name in the **Last Name** field.
2. The **Username** field updates automatically in real time, generating a value in the format `firstname.lastname` (all lowercase, without accented characters).
3. Click **Confirm**.

> The automatically generated username is permanent and cannot be changed later. It will be used to identify all RFQs and events you create.

The window cannot be closed without filling in both fields. The close button (X) has no effect on this screen.

---

## Splash Screen

After the identity is configured, a brief startup screen appears while the application initialises the database and working folders.

---

## Folder Structure Created Automatically

DataFlow automatically creates the following structure in the configured location (default: `Documents/DataFlow/`):

```
DataFlow/
├── Database/
│   └── dataflow_db.db          ← main database
└── Attachments/
    └── {RFQ number}/           ← attachments for each RFQ
```

Do not move these files manually while the application is open.

---

## Changing the Database Location

To use a shared network folder (for multi-user work):

1. Open **Settings** (⚙️ button in the toolbar).
2. In the **"DataFlow Standard Location"** section, read the current path.
3. Click **📁 Change DataFlow Location...** and choose the new folder.
4. Restart the application.

> After changing the location, the existing database is **not moved automatically**. Before restarting, manually copy the entire `DataFlow/` folder to the new location — otherwise the application will create a new, empty database.

See the [Multi-User Work](09-Multi-User-Work.md) section for full instructions.

---

## Your First Test RFQ

To get familiar with the application, here is how to create your first Request for Quotation in a few steps:

1. Click **➕ New Event** in the toolbar.
2. Select **📦 Full Supply**.
3. The RFQ control panel opens. Enter an expiry date and a reference (e.g. the project name).
4. Click **➕ Add Item** to insert the first line.
5. Click **Suppliers** to specify which suppliers to invite.
6. Close the window — the RFQ is saved automatically and will appear in the **Active RFQs** tab.

---

## Subsequent Launches

From the second launch onwards, DataFlow opens directly to the main screen without asking for your identity again. The application remembers the database location and the selected language.
