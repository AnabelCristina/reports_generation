# Excel Report Generator

Scripts em Python para gerar diferentes tipos de relatórios a partir de uma planilha Excel. Ideais para automatizar análise de KPIs, gráficos e exportação em PDF.

---

## 🚀 Funcionalidades

- Geração de KPIs limpos e estruturados.
- Criação de gráficos baseados nos dados da planilha.
- Exportação de relatórios em formato PDF.
- Possibilidade de segmentação por responsáveis.
- Geração de dados mock para testes.

---

## 📦 Requisitos

- Python 3.8 ou superior  
- Pacotes Python listados em `requirements.txt`, como por exemplo:
  - pandas
  - matplotlib
  - fpdf ou reportlab (para PDF)

---

## ⚙️ Instalação

1. **Clone este repositório:**
   ```bash
   git clone https://github.com/AnabelCristina/reports_generation.git
   cd reports_generation
   ```

2. **Crie e ative um ambiente virtual:**
   ```bash
   python -m venv venv
   source venv/bin/activate      # macOS/Linux
   venv\Scripts\activate         # Windows
   ```

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ Como Usar

**1. Gerar dados de exemplo**  
```bash
python create_mock_sheet_data.py
```
Gera uma planilha Excel de teste

**2. Criar um relatório de KPIs limpos**  
```bash
python generate_cleaned_kpi_report.py
```

**3. Gerar gráficos com base na planilha**  
```bash
python generate_graphics_report.py
```

**4. Exportar relatório consolidado em PDF**  
```bash
python generate_pdf_report.py
```

**5. Criar relatório por responsável**  
```bash
python generate_report_by_responsible.py
```
