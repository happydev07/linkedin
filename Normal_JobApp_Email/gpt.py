import requests
import json
import os
from docx import Document

def read_query_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read().strip()
    except Exception as e:
        print(f"Error reading the file: {e}")
        return None

def query_openai_direct(text):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {api_key}'
    }
    data = {
        "model": "gpt-4-turbo",
        "messages": [{"role": "system", "content": "You are a helpful assistant."},
                     {"role": "user", "content": text}]
    }
    
    response = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=data)
    if response.status_code == 200:
        return response.json()['choices'][0]['message']['content']
    else:
        raise Exception(f"API request failed with status {response.status_code}: {response.text}")

def delete_if_exists(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"Deleted existing file: {file_path}")

def copy_and_modify_docx(source_file, destination_file, new_text):
    if not os.path.exists(source_file):
        print(f"Source file does not exist: {source_file}")
        return

    doc = Document(source_file)
    for paragraph in doc.paragraphs:
        if '<>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<>', new_text)
    doc.save(destination_file)
    print(f"File copied and text replaced from {source_file} to {destination_file}")

# Specify the file path where the query is stored
query_file_path = 'C:\\Users\\devra\\Desktop\\BulkMail_JobApplication\\prompt.txt'

# Read the query from the file
query_text = read_query_from_file(query_file_path)

if query_text:
    # Get the response from OpenAI
    try:
        response_content = query_openai_direct(query_text)
        if response_content:
            # Perform file operations
            delete_if_exists('C:\\Users\\devra\\Desktop\\BulkMail_JobApplication\\AwsDevopsDeoraj.docx')
            copy_and_modify_docx('C:\\Users\\devra\\Desktop\\BulkMail_JobApplication\\deoraj.docx', 
                                 'C:\\Users\\devra\\Desktop\\BulkMail_JobApplication\\AwsDevopsDeoraj.docx', 
                                 response_content)
    except Exception as e:
        print(f"An error occurred while querying OpenAI: {e}")
else:
    print("Failed to read the query from the file.")
