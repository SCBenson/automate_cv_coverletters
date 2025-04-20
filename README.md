# Automate CV & Cover Letter Generator

A Python tool that automates the process of customizing your skills and qualifications based on job descriptions using Claude AI.

## Overview

This tool helps job seekers quickly create tailored skill lists and content for their CV/resume and cover letters by:

1. Taking your existing skills template and a job description
2. Merging them into a single document
3. Sending the merged document to Claude AI to extract relevant skills
4. Processing Claude's response to create a clean, structured list of skills and categories

## Features

- Merges a skills template with a job description
- Sends the merged document to Claude API for analysis
- Extracts structured lists from Claude's response
- Saves all outputs in a well-organized folder structure
- Formats results as clean, bullet-pointed lists in Word documents

## Requirements

- Python 3.6+
- An Anthropic API key for Claude AI

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/SCBenson/automate_cv_coverletters.git
   cd automate_cv_coverletters
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install python-docx anthropic python-dotenv
   ```

4. Create a `.env` file in the project root with your API key:
   ```
   ANTHROPIC_API_KEY=your_anthropic_api_key_here
   ```

## Project Structure

```
automate_cv_coverletters/
├── main.py                   # Main script
├── .env                      # Environment variables (API key)
├── assets/                   # Template files
│   ├── skills-prompt.docx    # Your skills template
│   └── job-description.docx  # Current job description
├── job_application_outputs/  # Generated outputs
│   ├── skills_prompt_w_job-*.docx    # Merged document
│   ├── claude_response-*.docx        # Claude's response
│   └── extracted_lists-*.docx        # Extracted skill lists
└── README.md                 # This file
```

## Usage

1. Put your skills template in `assets/skills-prompt.docx`
2. Put the job description in `assets/job-description.docx`
3. Run the script:
   ```bash
   python main.py
   ```

4. Check the `job_application_outputs` folder for the results

## How It Works

1. **Document Merging**: The tool takes your skills template and adds the job description to it
2. **Claude API**: The merged document is sent to Claude AI for analysis
3. **Response Processing**: Claude's response is extracted into clean, categorized skill lists
4. **Document Generation**: The extracted lists are formatted into a new Word document

## Customization

- Modify the main script to change the output folder name or file naming conventions
- Adjust the extraction patterns in `extract_list_items_text()` if needed for different formats
- Change the formatting in `save_extracted_lists()` to modify the output document style

## License

MIT

## Contributing

Feel free to submit issues or pull requests if you have suggestions for improvements!

## Acknowledgements

- [python-docx](https://python-docx.readthedocs.io/) for document handling
- [Anthropic Claude API](https://www.anthropic.com/claude) for AI-powered content analysis
