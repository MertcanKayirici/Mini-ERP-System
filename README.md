# 🚀 Mini ERP System (Excel VBA)

<p align="center">
  <img src="https://img.shields.io/badge/status-active-success" />
  <img src="https://img.shields.io/badge/platform-Excel%20VBA-blue" />
  <img src="https://img.shields.io/badge/architecture-layered-orange" />
</p>

---

> ⚡ A layered, failure-aware ERP simulation built in Excel VBA with transaction handling, retry mechanisms, and system monitoring — designed to reflect real-world backend architecture.

A structured mini ERP system that manages **documents, stock, ledger operations, and system state** with strong emphasis on **architecture, reliability, and error handling**.

---

## 🎬 Demo

<p align="center">
  <img src="assets/demo/test_engine_run.gif" width="100%" />
</p>

---

## 🚀 Core Capabilities

* Transaction-based document processing
* Automated stock & ledger integration
* Retry & recovery handling
* System-wide validation rules
* Test-driven verification system

---

## ✨ Key Features

* 🧱 Layered architecture (Service / Policy / Repository)
* 🔁 Retry & Recovery mechanism for failed operations
* ♻️ Idempotent processing (duplicate-safe operations)
* 🔒 Locking system with timeout handling
* 📊 Reconciliation system for data consistency
* 🧪 Automated test engine (PASS / FAIL)
* 📈 Dashboard for system monitoring

---

## 🧠 Architecture

![Architecture](docs/diagrams/system_architecture.png)

The system follows a clean layered structure:

* Service layer handles business logic
* Policy layer enforces rules
* Repository layer manages data
* Excel sheets act as storage

---

## 🔄 Transaction Flow

![Data Flow](docs/diagrams/data_flow.png)

A document operation flows through:

* Creation
* Validation
* Stock update
* Ledger entry
* Audit logging

---

## 🔁 Lifecycle

![Lifecycle](docs/diagrams/lifecycle.png)

Documents move through controlled states:

* Draft → Posting → Posted
* Cancel support
* RecoveryRequired state
* Retry mechanism

---

## 📊 Dashboard

![Dashboard](assets/screenshots/dashboard_overview.png)

The system monitor displays:

* Total stock
* Product count
* Ledger totals
* System status
* Last operations

---

## 🧪 Test System

| ✔️ All Tests Passing | ❌ Failure Handling |
|---------------------|-------------------|
| ![](assets/screenshots/test_results_all_pass.png) | ![](assets/screenshots/test_results_failure_case.png) |

The test engine validates:

* business rules
* system integrity
* error handling scenarios

---

## ⚙️ Tech Stack

* Excel VBA
* Layered Architecture Design
* Excel-based Data Storage
* Custom Test & Monitoring System

---

## 📂 Project Structure

```text
Mini-ERP-System/
├── docs/
├── src/
├── assets/
├── MiniERP_System.xlsm
└── README.md
```

---

## 🎯 Purpose

This project demonstrates:

* system design in constrained environments
* failure handling & retry logic
* layered architecture principles
* real-world backend simulation

---

## 💡 Why This Project?

This is not a simple CRUD system.

It focuses on:

* system behavior
* consistency control
* recoverability
* architectural thinking

---

## 👤 Developer

Mertcan Kayırıcı
Backend-Focused Developer

---
