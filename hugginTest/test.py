#Relevant metadata:
#Namn, dokumenttyp, process (3.4 medlemstöd), organisation (vem TAM har gjort dealen med), Klassificering (Publikt, internt), 
#Personuppgifter (inga, finns, känsliga), dokumentdatum, person som gav ut.

import torch
from transformers import pipeline
from shareplum import Site, Office365
from shareplum.site import Version
import docx

site_url = ''

documentText = ''

try:
    authcookie = Office365('', username='', password='').GetCookies()
    site = Site(site_url, version=Version.v365, authcookie=authcookie)

    folder = site.Folder('Delade dokument/2_Stod/2_9_IT')

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
