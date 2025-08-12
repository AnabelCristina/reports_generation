# Excel Report Generator

Python scripts that create different types of reports from an Excel sheet, that is mocked.


---

## 📦 Requirements

- Python 3.8 or higher
- Python packages (requirements.txt)
- - **wkhtmltopdf**:[Download e instalação](https://wkhtmltopdf.org/downloads.html)

---

## ⚙️ Instalation

1. **Clone:**
   ```bash
   git clone https://github.com/AnabelCristina/reports_generation.git
   cd reports_generation
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate      # macOS/Linux
   venv\Scripts\activate         # Windows
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ How to run

**1. Create mocked sheet**  
```bash
python create_mock_sheet_data.py
```
Gera uma planilha Excel de teste

**2. Generate graphics**  
```bash
python generate_graphics_report.py
```

**3. Export data to pdf**  
```bash
python generate_pdf_report.py
```

**5. Create reports by responsible**  
```bash
python generate_report_by_responsible.py
```
