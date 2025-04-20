from docx import Document
from docx.enum.text import WD_BREAK
import os
from datetime import datetime
from anthropic import Anthropic
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize Anthropic client
client = Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY"),
)

def create_output_folder(folder_name=None):

    if folder_name is None:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        folder_name = f"output_{timestamp}"

    output_folder = os.path.join(os.path.dirname(__file__), folder_name)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")
    else:
        print(f"Using existing output folder: {output_folder}")

    return output_folder

def merge_docx_files(skills_path, job_desc_path, output_folder, output_filename=None):
    """
    Loads two .docx files, creates a copy of the first one,
    and appends the second one to the first.
    """
    # Create output path if not provided
    if output_filename is None:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_filename = f"skills_prompt_w_job-{timestamp}.docx"
        # Save to the current directory, not in the assets folder
    output_path = os.path.join(output_folder, output_filename)

    print('Creating a copy of the skills prompt file')
    skills_doc = Document(skills_path)

    print('Loading the job description file.')
    job_desc_doc = Document(job_desc_path)

    skills_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    skills_doc.add_heading('Job Description', level=1)

    for paragraph in job_desc_doc.paragraphs: 
        new_paragraph = skills_doc.add_paragraph()
        for run in paragraph.runs:
            new_run = new_paragraph.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline

    skills_doc.save(output_path)
    print(f"Merged document saved to: {output_path}")
    
    return output_path

def extract_text_from_docx(docx_path):
    """
    Extract all text from a .docx file
    """
    doc = Document(docx_path)
    full_text = []
    
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    
    return "\n".join(full_text)

def ask_claude(prompt):
    """
    Send a prompt to Claude API and get the response
    """
    print("Sending prompt to Claude API...")
    
    message = client.messages.create(
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ],
        model="claude-3-5-sonnet-20240620",
    )
    
    return message.content[0].text

def save_response_to_docx(response_text, output_folder, output_filename=None):
    """
    Save Claude's response to a new .docx file
    """
    if output_filename is None:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_filename = f"claude_response-{timestamp}.docx"
    
    output_path = os.path.join(output_folder, output_filename)

    doc = Document()
    doc.add_heading('Claude Response', level=1)
    
    # Split the response by paragraphs and add each to the document
    paragraphs = response_text.split('\n')
    for para in paragraphs:
        if para.strip():  # Skip empty paragraphs
            doc.add_paragraph(para)
    
    doc.save(output_path)
    print(f"Claude's response saved to: {output_path}")
    
    return output_path

if __name__ == "__main__":

    # Create output folder.
    output_folder = create_output_folder("job_application_outputs")

    # Define paths relative to the script
    assets_folder = os.path.join(os.path.dirname(__file__), "assets")
    skills_path = os.path.join(assets_folder, "skills-prompt.docx")
    job_desc_path = os.path.join(assets_folder, "job-description.docx")
    
    # Verify files exist
    if not os.path.exists(skills_path):
        raise FileNotFoundError(f"Skills prompt file not found at {skills_path}")
    if not os.path.exists(job_desc_path):
        raise FileNotFoundError(f"Job description file not found at {job_desc_path}")
    
    # Step 1: Merge the documents
    merged_file = merge_docx_files(skills_path, job_desc_path, output_folder)
    print(f"Successfully created merged document: {merged_file}")
    
    # Step 2: Extract text from the merged document
    prompt_text = extract_text_from_docx(merged_file)
    print(f"Extracted {len(prompt_text)} characters from the merged document")
    
    # Step 3: Send to Claude API
    claude_response = ask_claude(prompt_text)
    print("Received response from Claude")
    
    # Step 4: Save Claude's response to a new document
    response_file = save_response_to_docx(claude_response, output_folder)
    print(f"Process completed. Response saved to: {response_file}")