# 💊 Inventory Pharmacy Management System

A desktop Pharmacy Inventory & Billing Management System built with **Visual Basic 6** and an **Access (ADO/JET)** database backend. Developed as an academic project to digitize day-to-day pharmacy operations — stock tracking, billing, and reporting — in a single offline desktop application.

---

## 📋 Overview

Managing a pharmacy manually means juggling stock registers, expiry dates, supplier records, and paper bills — all prone to human error. This system replaces that with a structured Windows desktop application where medicine stock, billing, and reports are handled through a connected Access database, with dedicated forms for each operation.

---

## ✨ Features

- **Inventory Management** — add, update, and track medicine stock with quantities and details
- **Billing** — generate customer bills against available stock
- **Database-Backed Records** — all data persisted via ADO to a local Access database, so nothing is lost between sessions
- **Reports** — generate stock and billing reports for record-keeping
- **Rich Form-Based UI** — dedicated VB6 forms for each core operation (inventory, billing, records)
- **Custom Icons & Resources** — polished desktop UI with dedicated icon and resource assets

---

## 🗂️ Project Structure

```
INVENTORY-PHARMACY/
├── Database/          # Access (.mdb) database file(s) — core data store
├── Documentation/      # Project report and supporting documentation
├── Forms/               # VB6 form files (.frm) — UI screens for the app
├── Icons/                # Icon assets used across the application
├── Modules/               # VB6 code modules (.bas) — shared logic & DB helpers
├── Project/                # VB6 project files (.vbp)
├── Reports/                 # Generated / template reports
├── Resources/                # Additional resource files used by the app
├── MSCOMCT2.OCX               # Microsoft Common Controls-2 (date/time picker, etc.)
├── RICHTX32.OCX                # Microsoft Rich Textbox control
└── compile.log                  # Build/compile log from the VB6 IDE
```

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Language | Visual Basic 6 (VB6) |
| Database | Microsoft Access (via ADO) |
| UI Controls | Microsoft Common Controls-2 (`MSCOMCT2.OCX`), Rich Textbox (`RICHTX32.OCX`) |
| Platform | Windows (32-bit desktop) |

---

## ⚙️ Requirements

- Windows OS (VB6 apps run natively on Windows; use a VM or compatibility layer on other OSes)
- Visual Basic 6.0 IDE (for editing/compiling from source) or the compiled `.exe` (for just running it)
- Microsoft Access Database Engine / JET OLEDB provider (for ADO connectivity to the `.mdb` file)
- The two `.OCX` controls (`MSCOMCT2.OCX`, `RICHTX32.OCX`) registered on the system — see setup below

---

## 🚀 Setup & Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/DharshanVelumani/INVENTORY-PHARMACY.git
   ```

2. **Register the required ActiveX controls** (run Command Prompt as Administrator)
   ```cmd
   regsvr32 MSCOMCT2.OCX
   regsvr32 RICHTX32.OCX
   ```

3. **Open the project**
   - Launch VB6 IDE
   - Open the `.vbp` project file from the `Project/` folder

4. **Verify the database path**
   - Ensure the Access database in `Database/` is referenced correctly in the connection string used by the app's modules

5. **Run**
   - Press `F5` in the VB6 IDE, or run the compiled `.exe` if one is included

---

## 📖 Documentation

Full project documentation, including the academic project report, is available in the [`Documentation/`](./Documentation) folder.

---

## 📌 Notes

- This project was built as part of a BCA academic curriculum to demonstrate database-driven desktop application development using VB6.
- Being a VB6/Access-based application, it targets Windows environments; it is not cross-platform.

---

## 📄 License

No license specified yet — all rights reserved by default until one is added.
