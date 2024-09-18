from flask import Flask, request, render_template, send_file
import os
import zipfile
from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf
import tabula
import pandas as pd
import logging
from io import BytesIO
import webbrowser
import threading
from PyPDF2 import PdfMerger

app = Flask(__name__)

# Cria a pasta de uploads, se não existir
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Obtém o diretório base do projeto
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')  # Define o caminho absoluto para a pasta de uploads
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Função para converter PDF para Word
def pdf_para_word(pdf_file, word_file):
    try:
        cv = Converter(pdf_file)
        cv.convert(word_file, start=0, end=None)
        cv.close()
        logging.debug(f"Arquivo convertido para Word: {word_file}")
    except Exception as e:
        logging.error(f"Erro ao converter PDF para Word: {e}")
        raise

# Função para converter Word para PDF
def word_para_pdf(word_file, output_pdf):
    try:
        docx_to_pdf(word_file, output_pdf)
        logging.debug(f"Arquivo convertido para PDF: {output_pdf}")
    except Exception as e:
        logging.error(f"Erro ao converter Word para PDF: {e}")
        raise

# Função para converter PDF para Excel
def pdf_para_excel(pdf_file, excel_file):
    try:
        tabula.convert_into(pdf_file, excel_file, output_format="csv", pages="all")
        logging.debug(f"Arquivo convertido para Excel: {excel_file}")
    except Exception as e:
        logging.error(f"Erro ao converter PDF para Excel: {e}")
        raise

# Função para converter Excel para PDF
def excel_para_pdf(excel_file, output_pdf):
    try:
        # Usando LibreOffice para conversão
        os.system(f'libreoffice --headless --convert-to pdf {excel_file} --outdir {os.path.dirname(output_pdf)}')
        logging.debug(f"Arquivo convertido para PDF: {output_pdf}")
    except Exception as e:
        logging.error(f"Erro ao converter Excel para PDF: {e}")
        raise

# Função para converter Excel para Word
def excel_para_word(excel_file, word_file):
    try:
        df = pd.read_excel(excel_file)
        from docx import Document
        doc = Document()

        for index, row in df.iterrows():
            doc.add_paragraph(str(row.to_dict()))  # Adiciona cada linha como um parágrafo
        doc.save(word_file)
        logging.debug(f"Arquivo convertido para Word: {word_file}")
    except Exception as e:
        logging.error(f"Erro ao converter Excel para Word: {e}")
        raise

# Função para juntar PDFs
def juntar_pdfs(pdf_files, output_pdf):
    try:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        merger.write(output_pdf)
        merger.close()
        logging.debug(f"Arquivos PDF juntados em: {output_pdf}")
    except Exception as e:
        logging.error(f"Erro ao juntar PDFs: {e}")
        raise

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    logging.debug("Iniciando o upload de arquivos.")

    if 'files' not in request.files:
        return "Nenhum arquivo foi selecionado.", 400

    files = request.files.getlist('files')
    format_to_convert = request.form.get('format')

    converted_files = []

    for file in files:
        if file.filename == '':
            continue

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        logging.debug(f"Arquivo salvo em: {file_path}")

        file_extension = os.path.splitext(file.filename)[1].lower()
        base_filename = os.path.splitext(file.filename)[0]

        if file_extension == '.pdf' and format_to_convert == 'word':
            converted_file = os.path.join(UPLOAD_FOLDER, base_filename + '.docx')
            pdf_para_word(file_path, converted_file)
            converted_files.append(converted_file)

        elif file_extension == '.pdf' and format_to_convert == 'excel':
            converted_file = os.path.join(UPLOAD_FOLDER, base_filename + '.csv')
            pdf_para_excel(file_path, converted_file)
            converted_files.append(converted_file)

        elif file_extension in ['.docx', '.doc'] and format_to_convert == 'pdf':
            converted_file = os.path.join(UPLOAD_FOLDER, base_filename + '.pdf')
            word_para_pdf(file_path, converted_file)
            converted_files.append(converted_file)

        elif file_extension in ['.xlsx', '.xls', '.csv'] and format_to_convert == 'pdf':
            converted_file = os.path.join(UPLOAD_FOLDER, base_filename + '.pdf')
            excel_para_pdf(file_path, converted_file)
            converted_files.append(converted_file)

        elif file_extension in ['.xlsx', '.xls', '.csv'] and format_to_convert == 'word':
            converted_file = os.path.join(UPLOAD_FOLDER, base_filename + '.docx')
            excel_para_word(file_path, converted_file)
            converted_files.append(converted_file)

        elif len(files) > 1 and format_to_convert == 'merge':
            output_pdf = os.path.join(UPLOAD_FOLDER, base_filename + '_merged.pdf')
            juntar_pdfs([file_path for file in files], output_pdf)
            converted_files.append(output_pdf)

        else:
            return "Formato de conversão não suportado ou arquivo inválido.", 400

    # Se houver apenas um arquivo convertido, envia diretamente
    if len(converted_files) == 1:
        try:
            return send_file(converted_files[0], as_attachment=True)
        except FileNotFoundError as e:
            logging.error(f"Arquivo não encontrado: {e}")
            return "Arquivo não encontrado, tente novamente.", 404

    # Se houver múltiplos arquivos, cria um zip em memória
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for file in converted_files:
            logging.debug(f"Adicionando arquivo ao ZIP: {file}")
            zipf.write(file, os.path.basename(file))

    zip_buffer.seek(0)  # Volta para o início do buffer

    logging.debug("Arquivo ZIP criado com sucesso.")
    return send_file(zip_buffer, as_attachment=True, download_name="converted_files.zip", mimetype='application/zip')

if __name__ == '__main__':
    # Usar um thread para executar o servidor Flask e abrir o navegador
    threading.Thread(target=lambda: webbrowser.open('http://127.0.0.1:5000')).start()
    app.run(debug=True, use_reloader=True)
