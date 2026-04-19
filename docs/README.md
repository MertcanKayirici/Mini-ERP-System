# 📚 Documentation

This directory contains the **architectural, technical, and system-level documentation** of the **Mini ERP System**.

It complements the source code by providing a clear explanation of **how the system is designed, structured, and operates internally**.

---

## 🏗️ Architecture Report

Detailed system design and architectural decisions are documented in:

* 📄 **DOCX Version**
  `architecture/Mini-ERP-Production-Architecture-Control-Report.docx`

* 📄 **PDF Version**
  `architecture/Mini-ERP-Production-Architecture-Control-Report.pdf`

---

### 📌 Report Covers

* System architecture overview
* Layered design (Service / Policy / Repository)
* Dependency rules and separation of concerns
* Transaction lifecycle and flow
* Retry & recovery mechanisms
* Idempotency strategy
* Reconciliation logic
* Design decisions and constraints

---

## 📊 System Diagrams

Visual representations of the system’s **structure and behavior**.

---

### 🧩 System Architecture

![System Architecture](diagrams/system_architecture.png)

Illustrates the layered architecture:

* **Service Layer** → business logic
* **Policy Layer** → validation and rules
* **Repository Layer** → data access
* **Excel Sheets** → data storage

---

### 🔄 Data Flow

![Data Flow](diagrams/data_flow.png)

Represents how data flows through the system:

* Document creation
* Validation pipeline
* Stock operations
* Ledger updates
* Audit logging

---

### 🔁 Document Lifecycle

![Lifecycle](diagrams/lifecycle.png)

Defines the state transitions of a document:

* Draft → Posting → Posted
* Cancel flow
* RecoveryRequired state
* Retry mechanism

---

## 🚀 Engineering Perspective

This documentation highlights:

* Separation of concerns through layered architecture
* Controlled state transitions
* Failure handling and recovery strategies
* Consistency validation via reconciliation
* System behavior beyond simple CRUD operations

---

## 🎯 Purpose

The goal of this documentation is to:

* Provide a clear understanding of system design
* Support technical evaluation (interviews, reviews)
* Explain internal workflows and decisions
* Demonstrate structured engineering thinking

---

## 🧠 Notes

* The system is implemented in **Excel VBA**
* Architecture follows a **layered design approach**
* Documentation is intended to complement the codebase
* Focus is on **system behavior, reliability, and structure** rather than UI

---

## 🔗 Related

Main project README: ../README.md

---
