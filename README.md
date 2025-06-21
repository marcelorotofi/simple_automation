# simple_automation

Python automation to split an Excel workbook into separate sheets per city, preserving original formatting.

---

## Table of Contents

- [Description](#description)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Input and Output](#input-and-output)
- [Customization](#customization)

---

## Description

This project provides a script (`simple_automation.py`) that:

1. Opens an Excel file (default: `popdata.xlsx`).
2. Reads the columns **Date of Birth**, **Full Name**, and **City** from the `Sheet1` tab.
3. For each unique city, creates a new sheet (if not already present).
4. Copies all rows for that city into the corresponding sheet, preserving styles and column widths.
5. Saves the result into a new file (`popdata2.xlsx`).

---

## Prerequisites

- Python 3.7 or higher
- [openpyxl](https://openpyxl.readthedocs.io/) library

Install dependencies with:

```bash
pip install openpyxl
Installation
Clone this repository:

bash
Copiar
Editar
git clone https://github.com/marcelorotofi/simple_automation.git
cd simple_automation
(Optional) Create and activate a virtual environment:

bash
Copiar
Editar
python -m venv .venv        # create venv
source .venv/bin/activate   # Linux/macOS
.venv\Scripts\activate      # Windows
Install dependencies:

bash
Copiar
Editar
pip install -r requirements.txt
Usage
Place your input file in the project root, named popdata.xlsx.

Run the script:

bash
Copiar
Editar
python src/simple_automation.py
At the end, a file named popdata2.xlsx will be generated, containing separate sheets for each city.

Project Structure
bash
Copiar
Editar
simple_automation/
├── src/
│   └── simple_automation.py   # Main script
├── tests/                     # (Future) unit tests
├── .gitignore                 # Ignores build artifacts and large files
├── requirements.txt           # Project dependencies
└── README.md                  # Project documentation
Input and Output
File	Description
popdata.xlsx	Original workbook with columns:
- Date of Birth
- Full Name
- City
popdata2.xlsx	Output workbook with separate sheets for each city

Customization
To change the input/output filenames, edit:

python
Copiar
Editar
workbook_city = load_workbook("popdata.xlsx")  # input file
...
workbook_city.save("popdata2.xlsx")            # output file
To modify the source sheet name, update:

python
Copiar
Editar
sheet_database = workbook_city['Sheet1']
