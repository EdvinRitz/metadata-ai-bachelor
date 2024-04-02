import torch
from transformers import AutoTokenizer, AutoModelForTokenClassification
from transformers import pipeline
from shareplum import Site, Office365
from shareplum.site import Version
import docx
from lxml import etree
import zipfile
import re
from memory_profiler import profile
import time


site_url = ''

documentText = ''

results_list = []  # This list will accumulate all results

#Token classification model
tokenizer = AutoTokenizer.from_pretrained("xlm-roberta-large-finetuned-conll03-english")
model = AutoModelForTokenClassification.from_pretrained("xlm-roberta-large-finetuned-conll03-english",
    ignore_mismatched_sizes=True  # Suppress warnings about mismatched sizes
    )
classifier = pipeline("ner", model=model, tokenizer=tokenizer)

#Question-answering model
qa_model = pipeline("question-answering", "timpal0l/mdeberta-v3-base-squad2")

#Questions
questionDocumentType = "Vad är titeln för dokumentet?"
questionProcess = "Vilken process beskriver dokumentet?"
questionPublisher = "Vem är ansvarig utgivare?"

#SharePlum information (Sensetive)
authcookie = Office365('', username='', password='').GetCookies()
site = Site(site_url, version=Version.v365, authcookie=authcookie)

def merge_and_clean_entities(entities, text):
    if not entities:
        return []

    # Sort entities by their start position
    entities.sort(key=lambda x: x['start'])

    merged_entities = []
    current_entity = entities[0].copy()  # Make a copy to avoid mutating the original

    for next_entity in entities[1:]:
        # Extract text between the current and next entity
        gap_text = text[current_entity['end']:next_entity['start']]
        
        # Lowercase comparison to decide on merging
        gap_text_lower = gap_text.lower()
        if gap_text_lower in [' ', ''] or re.match(r'^[\s,.]+$', gap_text_lower):
            # If entities are adjacent or separated by acceptable characters, merge them
            current_entity['end'] = next_entity['end']
        else:
            # Add the entity text (preserving original casing) before starting a new entity
            merged_entities.append(text[current_entity['start']:current_entity['end']])
            current_entity = next_entity.copy()

    # Don't forget the last entity
    merged_entities.append(text[current_entity['start']:current_entity['end']])

    # Clean and deduplicate while preserving original casing
    cleaned_entities = [clean_entity(entity) for entity in merged_entities]
    unique_entities_list = list(set(cleaned_entities))

    return unique_entities_list

# Define a clean_entity function that doesn't lower case but still removes unwanted characters
def clean_entity(entity):
    # First, strip leading/trailing whitespace characters, including newlines and tabs
    entity = entity.strip()

    # Then, remove leading/trailing punctuation using regex
    # \W matches any non-word character, equivalent to [^a-zA-Z0-9_], combined with strip() should cover the requirement
    entity = re.sub(r'^\W+|\W+$', '', entity)

    return entity

   
def process_docx_file(file_content, file_name):
    # Initialize a list to hold the results from this function
    local_results = []

    # Save the downloaded document
    with open(file_name, 'wb') as word_file:
        word_file.write(file_content)

    # Open and process the document
    doc = docx.Document(file_name)
    documentText = ""
    for para in doc.paragraphs:
        documentText += para.text + "\n"
    
    print('start')
    # Here you can do something with the extracted text
    print(f"Processed {file_name}")

    # Sample execution with your classifier and documentText
    resOrgs = classifier(documentText)

    # Assuming classifier results are stored in res1, and it includes 'start' and 'end' positions
    org_entities = [entity for entity in resOrgs if entity['entity'] in ('I-ORG', 'B-ORG')]

    # Preparing entities for merging, including word extraction
    for entity in org_entities:
        entity['word'] = documentText[entity['start']:entity['end']]

    # Merge, clean, and deduplicate entities
    merged_and_cleaned_org_entities = merge_and_clean_entities(org_entities, documentText)

    local_results.append('Organisations: ' + str(merged_and_cleaned_org_entities))
    print(merged_and_cleaned_org_entities)

    #These take alot of time
    resPublisher = qa_model(question = questionPublisher, context = documentText)
    resDocumentType = qa_model(question = questionDocumentType, context = documentText)
    resProcess = qa_model(question = questionProcess, context = documentText)

    #answer_DocumentType = resPublisher['answer']

    print('Ansvarig person är: ' + clean_entity(resPublisher['answer']))
    print('Dokumenttyp är: ' + clean_entity(resDocumentType['answer']))
    print('Processen är: ' + clean_entity(resProcess['answer']))

    local_results.append('Responsible publisher: ' + str(clean_entity(resPublisher['answer'])))
    local_results.append('Document type: ' + str(clean_entity(resDocumentType['answer'])))
    local_results.append('Process: ' + str(clean_entity(resProcess['answer'])))

    # Extract metadata
    with zipfile.ZipFile(file_name, 'r') as docx_zip:
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

            local_results.append(f"Creation date: {created.text if created is not None else 'Not available'}")
            local_results.append(f"Last modified date: {modified.text if modified is not None else 'Not available'}")
            local_results.append(f"Author: {creator.text if creator is not None else 'Not available'}")

        # Extracting application (program name) properties
        with docx_zip.open('docProps/app.xml') as app_xml:
            tree = etree.parse(app_xml)
            namespaces = {
                'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
            }
        
            application = tree.find('.//ep:Application', namespaces)
        
            print(f"Program Name: {application.text if application is not None else 'Not available'}")
            local_results.append(f"Program Name: {application.text if application is not None else 'Not available'}")
    print('end')
    # Return the list of results from this function
    return local_results
    
@profile #Comment this out if not testing
def explore_and_process_docx(folder_path, file_limit, results_list):
    # Check if the file processing limit has been reached
    if processed_files_counter['count'] >= file_limit:
        return  # Exit the function
    
    try:
        folder = site.Folder(folder_path)
        explored_folders_counter['count'] += 1
        # Process each file in the current folder
        for file_info in folder.files:
            file_name = file_info['Name']
            if file_name.endswith('.docx'): 
                print(f"Found .docx file: {file_name}")
                downloaded_doc = folder.get_file(file_name)
                
                processed_files_counter['count'] += 1

                #Editing endResult String for text file
                results_list.append('File number #' + str(processed_files_counter['count']))
                results_list.append('File name: ' + str(file_name))
                results_list.append('File path: ' + str(folder_path))

                #Proccessing file
                new_results = process_docx_file(downloaded_doc, file_name)
                results_list.extend(new_results)

                results_list.append('\n')
        
        # Attempt to list and explore subfolders
        subfolders = folder.folders
        for subfolder_name in subfolders:
            # Construct the path for the subfolder
            subfolder_path = f"{folder_path}/{subfolder_name}" if folder_path else subfolder_name
            print(f"Exploring subfolder: {subfolder_path}")
            explore_and_process_docx(subfolder_path, file_limit, results_list)
    except Exception as e:
        print(f"Error exploring {folder_path}: {e}")

# Start exploration from the root folder path
root_folder_path = 'Delade dokument'
processed_files_counter = {'count': 0}
explored_folders_counter = {'count': 0}

file_limit = 2 #Max number of files processed

start_time = time.time() #Comment out all time code if not testing

explore_and_process_docx(root_folder_path, file_limit, results_list)

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Total execution time: {elapsed_time} seconds")

print('Files processed: ' + str(processed_files_counter['count']))
print('Folders explored: ' + str(explored_folders_counter['count']))

# Construct the final result string
final_result_string = "\n".join(results_list)
#print(final_result_string)

#Upload results as a .txt file to the SharePoint
folder = site.Folder('Delade dokument')
#folder.upload_file(final_result_string, 'test.txt')
