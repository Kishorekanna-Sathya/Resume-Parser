import os
import re
import json
import pandas as pd
import google.generativeai as genai
from flask import Flask, jsonify
import fitz  
from docx import Document
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize Flask app
app = Flask(__name__)

# Load Gemini API Key
gemini_api_key = os.getenv('GEMINI_API_KEY')
if gemini_api_key:
    print("Gemini API Key loaded successfully.")
    genai.configure(api_key=gemini_api_key)
else:
    print("Gemini API Key not found!")
    exit(1)

# Directory containing resumes
RESUMES_DIR = os.getenv('RESUMES_DIR')

# Function to preprocess the extracted text
def preprocess_text(text):
    text = re.sub(r'\s+', ' ', text)  # Remove extra whitespace
    text = re.sub(r'[^\x00-\x7F]+', ' ', text)  # Remove non-ASCII characters
    return text.strip()

# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    try:
        doc = fitz.open(file_path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return ""

# Function to extract text from a DOCX file
def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return ""

# Function to process a single resume file
def process_resume(file_path):
    # Check the file type and process accordingly
    if file_path.lower().endswith('.pdf'):
        text = extract_text_from_pdf(file_path)
    elif file_path.lower().endswith('.docx'):
        text = extract_text_from_docx(file_path)
    else:
        print(f"Unsupported file type: {file_path}")
        return None

    if not text:
        return None

    # Preprocess the extracted text
    processed_text = preprocess_text(text)

    # Define the prompt for Gemini AI
    prompt = f"""
    Extract the following information from the resume in JSON format only:
    {{
        "Name": "<Name>",
        "Email": "<Email>",
        "Phone": "<Phone>",
        "College": "<College>",
        "City": "<City>",
        "Total Experience in years": "<Total Experience in years (include full time only and not part time/ internship)>",
        "Domain of Work": "<Domain of Work>",
        "Top Skills": "<Skill1>, <Skill2>, <Skill3>, <Skill4>, <Skill5>, <Skill6>, <Skill7>, <Skill8>, <Skill9>, <Skill10>",
        "Experience": "<CompanyName1 : Position Only>, 
                        <CompanyName2 : Position Only> 
        "Summary": "<Summary of candidate and experience>"
    }}
    
    Rules:
    1. If any field is unavailable, use 'NA'
    2. Only return valid JSON, no additional text
    3. Ensure all values are properly escaped for JSON
    4. The output must be parseable by json.loads()
    5. For Experience field, list companies in reverse chronological order (newest first)
    6. For Total Experience in years, calculate and provide a single number (e.g., "5.5")
    7. For Top Skills, provide exactly 10 skills as a comma-separated string
    8. For no experience then mention Experience as  Fresher in which the Total Experience in years should be as 0
    
    
    Resume text:
    {processed_text}
    """
    
    try:
        model = genai.GenerativeModel(model_name='gemini-1.5-flash')
        response = model.generate_content(prompt)
        
        # Try to extract JSON from the response
        response_text = response.text.strip()
        
        # Remove any markdown code blocks if present
        response_text = response_text.replace('```json', '').replace('```', '').strip()
        
        # Parse the JSON response
        extracted_details = json.loads(response_text)
        
        # Ensure all fields are present
        required_fields = [
            "Name", "Email", "Phone", "College", "City",
            "Total Experience in years", "Domain of Work", 
            "Top Skills", "Experience", "Summary"
        ]
        
        for field in required_fields:
            if field not in extracted_details:
                extracted_details[field] = "NA"
                
        return extracted_details
        
    except Exception as e:
        print(f"Error processing resume {file_path}: {e}")
        return None
@app.route('/')
def home():
    # Dictionary to store all resumes' data
    all_resumes_data = {}
    
    # Process all files in the resumes directory
    if os.path.exists(RESUMES_DIR):
        for filename in os.listdir(RESUMES_DIR):
            file_path = os.path.join(RESUMES_DIR, filename)
            if os.path.isfile(file_path):
                print(f"Processing: {filename}")
                details = process_resume(file_path)
                if details:
                    all_resumes_data[filename] = details
    
    # Save the extracted details to a JSON file and Excel file
    if all_resumes_data:
        # Save JSON
        with open('extracted_resume_details.json', 'w') as f:
            json.dump(all_resumes_data, f, indent=2)
        
        # Convert to DataFrame for Excel saving
        resumes_df = pd.DataFrame.from_dict(all_resumes_data, orient='index')
        
        # Save to Excel
        resumes_df.to_excel('extracted_resume_details.xlsx', index=True)
        
        # Prepare the data for HTML display
        html_content = "<html><body><h1>Extracted Resume Details</h1><ul>"
        
        for filename, details in all_resumes_data.items():
            html_content += f"<li><strong>{filename}</strong><ul>"
            for key, value in details.items():
                html_content += f"<li><strong>{key}:</strong> {value}</li>"
            html_content += "</ul></li>"
        
        html_content += "</ul></body></html>"
        
        # Return the HTML response with formatted details
        return html_content
    else:
        # Return a simple error message if no resumes were processed
        error_message = "<html><body><h1>Error: No resumes found or processed successfully.</h1></body></html>"
        return error_message


# Run the Flask app
if __name__ == '__main__':
    # Create resumes directory if it doesn't exist
    if not os.path.exists(RESUMES_DIR):
        os.makedirs(RESUMES_DIR)
        print(f"Created directory '{RESUMES_DIR}'. Please add your resume files there.")
    
    app.run(debug=True)