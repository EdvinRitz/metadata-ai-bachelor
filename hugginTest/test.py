#Relevant metadata:
#Namn, dokumenttyp, process (3.4 medlemstöd), organisation (vem TAM har gjort dealen med), Klassificering (Publikt, internt), 
#Personuppgifter (inga, finns, känsliga), dokumentdatum, person som gav ut.

import torch
from transformers import pipeline
from shareplum import Site, Office365
from shareplum.site import Version
import docx
from lxml import etree
import zipfile

site_url = ''

documentText = ''

modified = ''

try:
    authcookie = Office365('', username='', password='').GetCookies()
    site = Site(site_url, version=Version.v365, authcookie=authcookie)

    folder = site.Folder('')

    #for file_info in folder.files:
        #print(file_info['Name'])
        #print(file_info['ServerRelativeUrl'])

    #folder.upload_file('Test', 'test.txt')
    
    word_file_name = 'Webb-domäner-hosting-20190509.docx'

    downloaded_doc = folder.get_file(word_file_name)

    #print(download)

    with open(word_file_name, 'wb') as word_file:
        word_file.write(downloaded_doc)

    doc = docx.Document(word_file_name)

    for para in doc.paragraphs:
        documentText += para.text + "\n"
        #print(para.text)

    #print(documentText)
    # Extract metadata
    with zipfile.ZipFile(word_file_name, 'r') as docx_zip:
        # Extracting core properties including author and dates
        with docx_zip.open('docProps/core.xml') as core_xml:
            tree = etree.parse(core_xml)
            namespaces = {
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'dcmitype': 'http://purl.org/dc/dcmitype/',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
            }
        
            created = tree.find('.//dcterms:created', namespaces)
            modified = tree.find('.//dcterms:modified', namespaces)
            creator = tree.find('.//dc:creator', namespaces)
        
            print(f"Creation date: {created.text if created is not None else 'Not available'}")
            print(f"Last modified date: {modified.text if modified is not None else 'Not available'}")
            print(f"Author: {creator.text if creator is not None else 'Not available'}")

        # Extracting application (program name) properties
        with docx_zip.open('docProps/app.xml') as app_xml:
            tree = etree.parse(app_xml)
            namespaces = {
                'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
            }
        
            application = tree.find('.//ep:Application', namespaces)
        
            print(f"Program Name: {application.text if application is not None else 'Not available'}")

except Exception as e:
    print(f"An error occurred: {e}")


qa_model = pipeline("question-answering", "timpal0l/mdeberta-v3-base-squad2")

#"Vad är detta för dokument?" "Vad är detta för process?" "Vilket företag gäller dokumentet?" "Vad är dokumentdatumet?" "Vem är referensperson?" 

questionDocument = "Vem är ansvarig utgivare?"

context = documentText

res = qa_model(question = questionDocument, context = context)

answer_DocumentType = res['answer']

print(res)
#print(answer_DocumentType)
