import pandas as pd
import pdfkit
from jinja2 import Template
import os
import platform

# Configurar caminho do wkhtmltopdf conforme o sistema operacional

path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"  # Ajuste para o seu caminho!
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# 1. Ler o Excel com dados
df = pd.read_excel("team_kpis_mock.xlsx")

# 2. Template HTML para o relatório
html_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Report for {{ responsible }}</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2 { color: #2F5496; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px;}
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left;}
        th { background-color: #f2f2f2;}
    </style>
</head>
<body>
    <h2>Report for {{ responsible }}</h2>
    <table>
        <thead>
            <tr>
                {% for col in columns %}
                <th>{{ col }}</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in rows %}
            <tr>
                {% for col in columns %}
                <td>{{ row[col] }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
        </tbody>
    </table>
</body>
</html>
"""

# 3. Gerar PDFs individuais
responsibles = df["Responsible"].dropna().unique()

output_folder = "reports_pdfs"
os.makedirs(output_folder, exist_ok=True)

for responsible in responsibles:
    df_resp = df[df["Responsible"] == responsible].copy()
    df_resp = df_resp.drop(columns=["Responsible"])  # opcional, remover coluna Responsible do relatório

    template = Template(html_template)
    html_content = template.render(
        responsible=responsible,
        columns=df_resp.columns.tolist(),
        rows=df_resp.to_dict(orient="records")
    )

    # Salvar HTML temporário (opcional)
    temp_html = f"{output_folder}/temp_{responsible}.html"
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Converter HTML para PDF, passando a configuração do wkhtmltopdf
    pdf_path = f"{output_folder}/Report_{responsible}.pdf"
    if config:
        pdfkit.from_file(temp_html, pdf_path, configuration=config)
    else:
        pdfkit.from_file(temp_html, pdf_path)

    # Remover arquivo temporário
    os.remove(temp_html)

    print(f"PDF generated for {responsible}: {pdf_path}")

print("All PDFs generated!")