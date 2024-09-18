# Conversor de Arquivos

## Descrição

O Conversor de Arquivos é um aplicativo Flask que permite converter arquivos entre diferentes formatos, incluindo:

- PDF para Word (.docx)
- Word (.docx) para PDF
- PDF para Excel (.csv)
- Excel (.xls/.xlsx) para PDF

O aplicativo é fácil de usar e fornece uma interface web simples para upload e conversão de arquivos.

## Estrutura do Projeto


## Requisitos

Antes de executar o aplicativo, você precisará instalar as seguintes bibliotecas Python:

- Flask
- pdf2docx
- docx2pdf
- tabula-py
- pandas
- PyInstaller (opcional, para criar executável)

1. **Você pode instalar as dependências necessárias usando o pip:**
    ```bash
    pip install Flask pdf2docx docx2pdf tabula-py pandas


# Criando um Executável

## Se você deseja criar um executável do aplicativo, siga os passos abaixo:

1. **Instale o PyInstaller (caso ainda não tenha feito):**
    ```bash
    pip install pyinstaller

2.  **Para criar um executavel**
    ```bash
    pyinstaller --onefile --add-data "templates:templates" --hidden-import pdf2docx --hidden-import docx2pdf conv.py

3.  **Caso não funcione**
    ```bash
    pip install --upgrade pyinstaller

4. **E faça novamente**
    ```bash
    pyinstaller --onefile --add-data "templates:templates" --add-data "uploads:uploads" --hidden-import pdf2docx --hidden-import docx2pdf --hidden-import pandas --hidden-import tabula --hidden-import PyPDF2 conv.py

### Notas Adicionais    

- **Personalize conforme necessário**: Adapte a descrição, estrutura do projeto e instruções de execução de acordo com os detalhes do seu aplicativo.
- **Adicionar mais detalhes**: Você pode incluir informações sobre como usar a interface web, exemplos de uso, ou quaisquer dependências específicas que você esteja utilizando.




