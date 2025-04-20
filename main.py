from docx import Document
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import re
import shutil
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
    """Create an output folder for saving files."""
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

def extract_skills_from_response(docx_path):
    """
    Extract main skills and sub-skills from Claude's response.
    """
    text = extract_text_from_docx(docx_path)

    lines = [line.strip() for line in text.split('\n') if line.strip()]

    if lines and lines[0].startswith('Claude Response'):
        lines = lines[1:]
    
    main_item_pattern = re.compile(r'^(\d+)\)\s+(.*)$')
    sub_item_pattern = re.compile(r'^\(([ivxl]+)\)\s+(.*)$')

    main_items_text = []
    sub_items_text = []

    for line in lines:
        main_match = main_item_pattern.match(line)
        if main_match:
            text = main_match.group(2)
            main_items_text.append(text)
            continue

        sub_match = sub_item_pattern.match(line)
        if sub_match:
            text = sub_match.group(2)
            sub_items_text.append(text)
    
    return main_items_text, sub_items_text

def update_skills_table(cv_template_path, main_skills, sub_skills, output_folder, output_filename=None):
    """
    Updates the existing empty skills table in the CV template with the provided skills,
    with all text left-aligned.
    
    Parameters:
    cv_template_path (str): Path to the CV template document with existing table
    main_skills (list): List of main skills
    sub_skills (list): List of sub-skills (5 sub-skills per main skill)
    output_folder (str): Folder to save the output document
    output_filename (str, optional): Name for the output file
    
    Returns:
    str: Path to the saved document
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import os
    from datetime import datetime
    import shutil
    
    # Create output filename if not provided
    if output_filename is None:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_filename = f"CV_with_skills-{timestamp}.docx"
    
    output_path = os.path.join(output_folder, output_filename)
    
    # Make a copy of the template first to avoid modifying the original
    temp_path = os.path.join(output_folder, "temp_" + os.path.basename(cv_template_path))
    shutil.copy2(cv_template_path, temp_path)
    
    # Load the CV template copy
    doc = Document(temp_path)
    
    # Find the skills table (first table in the document)
    if len(doc.tables) == 0:
        print("Warning: No tables found in the template.")
        os.remove(temp_path)  # Clean up
        return None
    
    # Get the first table
    skills_table = doc.tables[0]
    
    # Make sure we have the right number of sub-skills (5 per main skill)
    if len(sub_skills) < len(main_skills) * 5:
        print(f"Warning: Not enough sub-skills provided. Expected {len(main_skills) * 5}, got {len(sub_skills)}")
        # Pad with empty strings if necessary
        sub_skills.extend([''] * (len(main_skills) * 5 - len(sub_skills)))
    
    # Calculate how many skill pairs we need (2 columns per row)
    num_pairs = (len(main_skills) + 1) // 2  # Ceiling division to handle odd number of main skills
    
    # Add more rows if needed
    while len(skills_table.rows) < num_pairs:
        skills_table.add_row()  # Add more rows if needed
    
    # Fill the table with skills
    for pair_idx in range(num_pairs):
        # Calculate indices for this pair
        left_skill_idx = pair_idx * 2
        right_skill_idx = pair_idx * 2 + 1
        
        # Get the current row
        row = skills_table.rows[pair_idx]
        
        # Update left cell
        if left_skill_idx < len(main_skills):
            left_cell = row.cells[0]
            
            # Check if the cell has any paragraphs
            if len(left_cell.paragraphs) == 0:
                left_cell.add_paragraph()
            
            # Clear any existing content
            for i in range(len(left_cell.paragraphs)):
                if i > 0:  # Keep the first paragraph, remove extras
                    left_cell._element.remove(left_cell.paragraphs[i]._p)
            
            # Now we have just one empty paragraph
            p = left_cell.paragraphs[0]
            p.text = ""  # Ensure it's empty
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Left alignment
            run = p.add_run(main_skills[left_skill_idx])
            run.bold = True
            
            # Add sub-skills for left skill
            for i in range(5):
                sub_idx = left_skill_idx * 5 + i
                if sub_idx < len(sub_skills) and sub_skills[sub_idx]:
                    p = left_cell.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Left alignment
                    p.add_run(sub_skills[sub_idx])
        
        # Update right cell
        if right_skill_idx < len(main_skills):
            right_cell = row.cells[1]
            
            # Check if the cell has any paragraphs
            if len(right_cell.paragraphs) == 0:
                right_cell.add_paragraph()
            
            # Clear any existing content
            for i in range(len(right_cell.paragraphs)):
                if i > 0:  # Keep the first paragraph, remove extras
                    right_cell._element.remove(right_cell.paragraphs[i]._p)
            
            # Now we have just one empty paragraph
            p = right_cell.paragraphs[0]
            p.text = ""  # Ensure it's empty
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Left alignment
            run = p.add_run(main_skills[right_skill_idx])
            run.bold = True
            
            # Add sub-skills for right skill
            for i in range(5):
                sub_idx = right_skill_idx * 5 + i
                if sub_idx < len(sub_skills) and sub_skills[sub_idx]:
                    p = right_cell.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Left alignment
                    p.add_run(sub_skills[sub_idx])
    
    # Save the document
    doc.save(output_path)
    print(f"CV with updated skills saved to: {output_path}")
    
    # Clean up the temporary file
    os.remove(temp_path)
    
    return output_path

if __name__ == "__main__":
    # Create output folder
    output_folder = create_output_folder("job_application_outputs")

    # Define paths relative to the script
    assets_folder = os.path.join(os.path.dirname(__file__), "assets")
    skills_path = os.path.join(assets_folder, "skills-prompt.docx")
    job_desc_path = os.path.join(assets_folder, "job-description.docx")
    cv_template_path = os.path.join(assets_folder, "CV_Template.docx")
    
    # Verify files exist
    if not os.path.exists(skills_path):
        raise FileNotFoundError(f"Skills prompt file not found at {skills_path}")
    if not os.path.exists(job_desc_path):
        raise FileNotFoundError(f"Job description file not found at {job_desc_path}")
    if not os.path.exists(cv_template_path):
        raise FileNotFoundError(f"CV template file not found at {cv_template_path}")
    
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
    print(f"Response saved to: {response_file}")

    # Step 5: Extract skills from Claude's response
    main_items, sub_items = extract_skills_from_response(response_file)
    print(f"Extracted {len(main_items)} main skills and {len(sub_items)} sub-skills")
    
    # Step 6: Update the skills table in the CV template
    updated_cv = update_skills_table(
        cv_template_path, 
        main_items,  # Your main skills list
        sub_items,   # Your sub-skills list
        output_folder
    )
    print(f"Updated CV saved to: {updated_cv}")