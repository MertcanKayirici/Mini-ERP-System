# 🚀 Mini ERP System (Excel VBA)

A **layered, failure-aware Mini ERP system** built with **Excel VBA**, designed to simulate real-world backend architecture including transaction handling, retry mechanisms, and system monitoring.

---

## 🎯 Key Features

* 🧱 **Layered Architecture**

  * Service / Policy / Repository separation
* 🔁 **Retry & Recovery System**

  * Handles partial failures and resumes safely
* ♻️ **Idempotent Operations**

  * Prevents duplicate processing
* 🔒 **Locking Mechanism**

  * Prevents concurrent conflicts with timeout handling
* 📊 **Reconciliation Engine**

  * Ensures system data consistency
* 🧪 **Automated Test Engine**

  * PASS / FAIL validation with real business errors
* 📈 **System Monitoring Dashboard**

  * Live system status and metrics

---

## 🧠 Architecture Overview

![Architecture](docs/diagrams/system_architecture.png)

* **Service Layer** → business logic
* **Policy Layer** → validation & rules
* **Repository Layer** → data access
* **Excel Sheets** → data storage

---

## 🔄 Transaction Flow

![Data Flow](docs/diagrams/data_flow.png)

* Document creation
* Validation
* Stock operations
* Ledger updates
* Audit logging

---

## 🔁 Document Lifecycle

![Lifecycle](docs/diagrams/lifecycle.png)

* Draft → Posting → Posted
* Cancel support
* RecoveryRequired state
* Retry mechanism

---

## 🧪 Test Engine (Live Demo)

![Test Engine](assets/demo/test_engine_run.gif)

* Automated test execution
* PASS / FAIL results
* Real error handling (e.g., insufficient stock, inactive product)

---

## 📊 Dashboard

![Dashboard](assets/screenshots/dashboard_overview.png)

Displays:

* Total stock
* Product count
* Ledger total
* Last operation
* System status
* Test results summary

---

## 📸 Test Results

### ✔️ Successful Run

![All Pass](assets/screenshots/test_results_all_pass.png)

### ❌ Failure Handling

![Failure](assets/screenshots/test_results_failure_case.png)

---

## 📁 Project Structure

```text
Mini-ERP-System/
├── docs/
├── src/
├── assets/
├── database/
├── MiniERP_System.xlsm
└── README.md
```

---

## ⚙️ Tech Stack

* **Excel VBA**
* Layered Architecture Design
* Manual Data Storage (Excel Sheets)
* Custom Test & Monitoring System

---

## 📌 Notes

* This project focuses on **system design and reliability**, not UI
* VBA code is exported as `.bas` and `.cls` files for version control
* Designed to demonstrate **engineering thinking beyond CRUD applications**

---

## 🎯 Why This Project?

This project demonstrates:

* Real-world system behavior simulation
* Error handling & recovery design
* Clean architecture principles in a constrained environment (Excel VBA)

---

## 👤 Author

Developed as a portfolio project to showcase backend/system design skills using Excel VBA.

---
