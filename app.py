import os
from flask import Flask, render_template, request, send_file
from flask_mysqldb import MySQL
import pandas as pd
import pdfkit
from datetime import datetime

app = Flask(__name__)

# Configurações do MySQL
app.config['MYSQL_HOST'] = '127.0.0.1'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'diario_obra_mysql'

# Configuração do MySQL
mysql = MySQL(app)

# Diretórios para arquivos Excel e PDF
base_directory = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIRECTORY = os.path.join(base_directory, 'excel_files')
PDF_DIRECTORY = os.path.join(base_directory, 'pdf_files')

# Verifica e cria os diretórios se não existirem
if not os.path.exists(EXCEL_DIRECTORY):
    os.makedirs(EXCEL_DIRECTORY)

if not os.path.exists(PDF_DIRECTORY):
    os.makedirs(PDF_DIRECTORY)

# Configura os diretórios
app.config['EXCEL_DIRECTORY'] = EXCEL_DIRECTORY
app.config['PDF_DIRECTORY'] = PDF_DIRECTORY


# Função para criar o arquivo Excel para uma entrada específica
def create_excel_file(entry_id):
    conn = mysql.connection
    cursor = conn.cursor()

    # Consulta para obter uma entrada específica no MySQL
    cursor.execute("SELECT * FROM entries WHERE id = %s", (entry_id,))
    data = cursor.fetchall()

    if data:
        # Cria um DataFrame pandas
        df = pd.DataFrame(data, columns=[col[0] for col in cursor.description])

        # Verifica e cria o diretório de arquivos Excel se não existir
        if not os.path.exists(app.config['EXCEL_DIRECTORY']):
            os.makedirs(app.config['EXCEL_DIRECTORY'])

        # Nome do arquivo Excel baseado no ID da entrada e no timestamp atual
        excel_file_path = os.path.join(app.config['EXCEL_DIRECTORY'], f'entry_{entry_id}_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
        df.to_excel(excel_file_path, index=False)

        return excel_file_path
    else:
        return None


# Função para criar o arquivo PDF para uma entrada específica
def create_pdf_file(entry_id):
    conn = mysql.connection
    cursor = conn.cursor()

    # Consulta para obter uma entrada específica no MySQL
    cursor.execute("SELECT * FROM entries WHERE id = %s", (entry_id,))
    data = cursor.fetchall()

    if data:
        html_content = "<html><head><style>"
        # Estilos CSS
        html_content += """
            body {
                font-family: Arial, sans-serif;
                background-color: #f4f4f4;
                margin: 0;
                padding: 0;
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
            }
            th {
                background-color: #f2f2f2;
            }
            /* Estilos adicionais aqui... */
            .red-text {
                color: red;
            }
        """
        html_content += "</style></head><body><table>"

        # Criação do conteúdo do PDF com todos os campos em uma tabela
        for index, entry in enumerate(data):
            html_content += "<tr>"
            html_content += f"<th>Field</th><th>Value</th>"
            html_content += "</tr>"

            fields = ["Title", "Date", "Location", "Description", "Cost", "Status", "Activity Type", "Responsible", "Priority", "URL", "Attachment"]
            for i in range(len(fields)):
                html_content += "<tr>"
                html_content += f"<td>{fields[i]}</td>"
                # Adiciona a classe red-text para tornar o texto vermelho
                html_content += f"<td class='red-text'>{entry[i+1]}</td>"
                html_content += "</tr>"

            html_content += "</table></body></html>"

            html_file_path = os.path.join(app.config['PDF_DIRECTORY'], f'entry_{entry_id}_{datetime.now().strftime("%Y%m%d%H%M%S")}.html')

            if not os.path.exists(app.config['PDF_DIRECTORY']):
                os.makedirs(app.config['PDF_DIRECTORY'])

            with open(html_file_path, 'w') as html_file:
                html_file.write(html_content)

            pdf_file_path = os.path.join(app.config['PDF_DIRECTORY'], f'entry_{entry_id}_{datetime.now().strftime("%Y%m%d%H%M%S")}.pdf')

            config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

            pdfkit.from_file(html_file_path, pdf_file_path, configuration=config)

            return True

    else:
        return False



# Rota principal para exibir arquivos Excel e PDF disponíveis
@app.route('/')
def show_entries():
    # Obtém os arquivos disponíveis no diretório de arquivos Excel
    excel_files = os.listdir(app.config['EXCEL_DIRECTORY'])
    excel_files = [file for file in excel_files if file.endswith('.xlsx')]
    
    # Obtém os arquivos disponíveis no diretório de arquivos PDF
    pdf_files = os.listdir(app.config['PDF_DIRECTORY'])
    pdf_files = [file for file in pdf_files if file.endswith('.pdf')]
    
    # Cria uma lista de tuplas para os arquivos Excel contendo o nome e o link de download
    excel_entries = [(file, f'/download_file/excel/{file}') for file in excel_files]

    # Cria uma lista de tuplas para os arquivos PDF contendo o nome e o link de download
    pdf_entries = [(file, f'/download_file/pdf/{file}') for file in pdf_files]
    
    return render_template('index.html', excel_entries=excel_entries, pdf_entries=pdf_entries)


# Rota para adicionar entrada
@app.route('/add_entry', methods=['POST'])
def add_entry():
    if request.method == 'POST':
        # Código para inserir dados no banco de dados MySQL
        title = request.form['title']
        date = request.form['date']
        location = request.form['location']
        description = request.form['description']
        cost = request.form['cost']
        status = request.form['status']
        activity_type = request.form['activity_type']
        responsible = request.form['responsible']
        priority = request.form['priority']
        url = request.form['url']
        attachment = request.form['attachment']
        
        # Insere os dados no banco de dados MySQL
        conn = mysql.connection
        cursor = conn.cursor()
        cursor.execute("INSERT INTO entries (title, date, location, description, cost, status, activity_type, responsible, priority, url, attachment) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                       (title, date, location, description, cost, status, activity_type, responsible, priority, url, attachment))
        conn.commit()
        cursor.close()

        # Obtém o ID da última entrada adicionada
        cursor = conn.cursor()
        cursor.execute("SELECT LAST_INSERT_ID()")
        entry_id = cursor.fetchone()[0]
        cursor.close()

        # Cria um novo arquivo Excel para a entrada adicionada
        create_excel_file(entry_id)

        # Cria um novo arquivo PDF para a entrada adicionada
        create_pdf_file(entry_id)

        return 'Entrada adicionada com sucesso'


# Rota para baixar arquivo Excel
@app.route('/download_file/excel/<file_name>')
def download_excel_file(file_name):
    # Verifica se o arquivo Excel está no diretório e envia para download
    excel_file_path = os.path.join(app.config['EXCEL_DIRECTORY'], file_name)
    
    if os.path.exists(excel_file_path):
        return send_file(excel_file_path, as_attachment=True)
    else:
        return "Arquivo Excel não encontrado"


# Rota para baixar arquivo PDF
@app.route('/download_file/pdf/<file_name>')
def download_pdf_file(file_name):
    # Verifica se o arquivo PDF está no diretório e envia para download
    pdf_file_path = os.path.join(app.config['PDF_DIRECTORY'], file_name)
    
    if os.path.exists(pdf_file_path):
        return send_file(pdf_file_path, as_attachment=True)
    else:
        return "Arquivo PDF não encontrado"


if __name__ == '__main__':    
    app.run(debug=True)
