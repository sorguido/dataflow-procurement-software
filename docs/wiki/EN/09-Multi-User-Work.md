# 09 – Multi-User Work

## How Database Sharing Works

DataFlow supports concurrent use by multiple users on the same database. The mechanism relies on a shared SQLite file in a network folder (company server, NAS, or shared drive).

**No application server is required.** Each user runs their own installation of DataFlow on their own computer; everyone points to the same network folder where the database resides.

---

## Setting Up the Shared Folder

### Step 1 – Prepare the Folder on the Server

On a network path accessible to all users (e.g. `\\server\DataFlow\`), create the following structure:

```
DataFlow\
├── Database\
└── Attachments\
```

If a user's database already exists, copy the `dataflow_db.db` file into `DataFlow\Database\`.

### Step 2 – Configure Each Workstation

On each computer:

1. Open DataFlow.
2. Go to **⚙️ Settings**.
3. Click **📁 Change DataFlow Location...** and select the network folder (e.g. `\\server\DataFlow\`).
4. Restart DataFlow.

### Step 3 – Verify User Identity

Each user must have a unique identity (the username generated at first launch). DataFlow uses the username to distinguish each buyer's data. If two users share the same name (e.g. two people named "John Smith"), they may end up with identical usernames; in that case, one of them must be renamed before sharing the database (this requires technical assistance).

---

## How DataFlow Handles the Shared Database

- The database uses SQLite's **WAL (Write-Ahead Log)** mode, which allows **multiple simultaneous readers** and **one writer at a time**.
- Writes have a timeout of **10 seconds**: if the database is busy with another write, DataFlow waits up to 10 seconds before reporting an error.
- Read operations (browsing, searching) use **read-only** access so they never interfere with ongoing writes.

Under normal conditions with 5–10 users, no conflicts will be visible.

---

## Data Visibility

Every user can see **all data from all users**, subject to the following restrictions:

| Data | Owner | Other user |
|------|-------|------------|
| RFQ | Opens and edits freely | Opens in read-only mode |
| Value Stream Mapping event (Savings/CA) | Edits and deletes | Read-only |
| Derisking supplier | Edits and deletes | Read-only |

**Read-only mode** is indicated by:
- A red banner at the bottom of the RFQ window: *"⚠️ READ-ONLY MODE: You are viewing an RFQ belonging to another user."*
- All edit fields and buttons disabled.
- Opening and downloading attachments remain available.

---

## Filtering by User

In the advanced filters panel, the **User** field lets you filter to see only one specific buyer's RFQs, or all of them at once.

- Select your own name to see only your RFQs (searches the local database only).
- Select **"(All Users)"** to see all buyers' RFQs (aggregated search across all databases in the same `Database/` folder).
- Select a colleague's name to see only their RFQs.

---

## Multiple Databases in the Same Folder

DataFlow also supports an advanced sharing variant: each user can have their **own separate database file** inside the same `Database/` folder. In this case, the **aggregated search** function reads all `*.db` files in the folder and presents them to the user as if they were one (other users' data is read-only).

This mode activates automatically when each user points to the same network folder but created their own database locally and then copied it there.

---

## Backup and Shared Database

When the database is shared on a network, the daily automatic backup should be configured on **one workstation only** (e.g. the purchasing manager's computer or a server). Configuring it on multiple workstations creates redundant backups, which is acceptable but unnecessary.

Before performing a manual backup of a shared database, verify that no other user is currently performing write operations.

---

## Behaviour When the Network Is Unavailable

If the network folder cannot be reached at startup, DataFlow cannot find the database and displays an error on opening. In this case:

1. Check the network connection.
2. Ensure you have read/write permissions on the shared folder.
3. If necessary, work temporarily with a local copy of the database and re-align it manually afterwards.

DataFlow does not automatically handle merge conflicts between two databases that have evolved independently.
