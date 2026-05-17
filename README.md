# 📄 ML Experiment Report Generator

> A simple script that automatically compiles ML experiment outputs — images, data tables, and logs — into a single Word document.

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python) ![python-docx](https://img.shields.io/badge/Library-python--docx-green) ![pandas](https://img.shields.io/badge/Library-pandas-orange?logo=pandas) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)

---

## 📌 Overview

When running multiple ML experiments, results end up scattered across dozens of folders — confusion matrices, ROC curves, training logs, model summaries. Reviewing them manually is tedious.

This script walks through all experiment folders and compiles everything into one clean Word document, organized by case and set — in one run.

---

## ✨ Features

| Feature | Description |
|--------|-------------|
| 📁 Auto folder scan | Detects all subfolders (e.g. A1, A2, B1, B2...) automatically |
| 🖼️ All images included | Pulls confusion matrix, ROC curve, training curves — all in order |
| 📊 Full data tables | All CSV / Excel files converted to Word tables (no row limit) |
| 📝 Text file support | Model architecture and notes included as-is |
| 🗂️ Case separation | Generates separate reports for Case A and Case B |
| 📄 Auto page breaks | Each experiment folder starts on a new page |

---

## 🖥️ How to Use

**1. Install dependencies**
```bash
pip install python-docx pandas openpyxl
```

**2. Set your project path in the script**
```python
BASE_PATH = r"your\project\folder\path"
```

**3. Run**
```bash
python generate_report.py
```

**4. Output files are saved in your project root**
```
Full_Report_CaseA.docx
Full_Report_CaseB.docx
```

---

## 📂 Expected Folder Structure

```
my_project/
├── models_set1_caseA/
│   ├── A1/
│   │   ├── confusion_matrix.png
│   │   ├── roc_curve.png
│   │   ├── train_log.csv
│   │   └── model_architecture.txt
│   ├── A2/
│   └── ...
├── models_set2_caseA/
├── models_set1_caseB/
└── ...
```

---

## 💡 Background

Built as a helper tool during a capstone project on motor fault diagnosis using the NASA IMS bearing dataset. After each experiment run, results were spread across 20+ folders — this script was written to make reviewing them significantly faster.

---

## 🛠️ Tech Stack

- **Language:** Python 3
- **Libraries:** python-docx, pandas, openpyxl
