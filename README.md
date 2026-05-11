# CNDAG Word Sub‑Document Generator

## Overview

This project generates **exact Word sub‑documents** from a reference `.docx` file using a Base64‑encoded JSON configuration.  
Each sub‑document corresponds to a section tagged with **`CNDAG PIPELINE`** and preserves **word‑to‑word formatting fidelity**.

The original reference document is **never modified**.

---

## What This Solves

Given:
- A Base64‑encoded JSON input
- Sections tagged with `dataSource.type = "CNDAG PIPELINE"`
- A reference Word document containing all sections

The system:
- Identifies all CNDAG‑tagged sections
- Produces **one `.docx` per section**
- Preserves **exact formatting**, including:
  - Tables and borders
  - Bullets and numbering
  - Indentation and spacing
  - Code blocks (bash/CLI)
  - Diagrams and images

---

## Core Design Principle

> **Exact Word fidelity is achieved by subtraction, not reconstruction.**

For each CNDAG section:
1. Copy the entire reference Word document
2. Remove all content except the target section
3. Save the result as a new sub‑document

This keeps all document‑level styles, themes, and numbering intact.

---

## Workflow

