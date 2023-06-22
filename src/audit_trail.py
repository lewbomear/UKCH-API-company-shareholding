import pdfkit
#needs system install of wkhtmltopdf: https://wkhtmltopdf.org/downloads.html

def save_page_as_pdf(pdf_url, pdf_path, company_name):
    options = {
        'quiet': '',
        'page-size': 'A4',
        'margin-top': '10mm',
        'margin-right': '10mm',
        'margin-bottom': '10mm',
        'margin-left': '10mm',
    }
    pdfkit.from_url(pdf_url, pdf_path.format(company_name), options=options)

