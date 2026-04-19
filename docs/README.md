# 📚 Documentation

This folder contains the architectural and technical documentation of the **Mini ERP System**.

---

## 🏗️ Architecture Report

Detailed system design and architectural decisions are documented in the following files:

- 📄 **DOCX Version**  
  `architecture/Mini-ERP-Production-Architecture-Control-Report.docx`

- 📄 **PDF Version**  
  `architecture/Mini-ERP-Production-Architecture-Control-Report.pdf`

### 📌 Content Includes:
- System architecture overview
- Layered design (Service / Policy / Repository)
- Dependency rules and boundaries
- Transaction management
- Retry & recovery mechanisms
- Idempotency strategy
- Reconciliation logic
- System constraints and decisions

---

## 📊 Diagrams

Visual representations of the system architecture and behavior.

---

### 🧩 System Architecture

![System Architecture](diagrams/system_architecture.png)

Represents the layered architecture of the system:

- Service Layer (business logic)
- Policy Layer (validation & rules)
- Repository Layer (data access)
- Excel-based data storage

---

### 🔄 Data Flow

![Data Flow](diagrams/data_flow.png)

Illustrates how data flows through the system:

- Document creation
- Validation
- Stock operations
- Ledger updates
- Audit logging

---

### 🔁 Document Lifecycle

![Lifecycle](diagrams/lifecycle.png)

Shows the state transitions of a document:

- Draft → Posting → Posted
- Cancel flow
- RecoveryRequired state
- Retry mechanism

---

## 🎯 Purpose

This documentation is designed to:

- Explain the internal architecture clearly
- Provide a quick understanding of system behavior
- Support technical evaluation (e.g., interviews, reviews)
- Demonstrate engineering thinking beyond basic CRUD systems

---

## 🧠 Notes

- The system is implemented in **Excel VBA**
- Source code is structured using layered architecture principles
- Documentation complements the codebase for better understanding

---