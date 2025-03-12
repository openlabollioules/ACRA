# ACRA : Automatic CRA Generator

## Overview

Automatic CRA Generator is a tool designed to aggregate and summarize multiple Activity Reports (CRA - *Compte Rendu d'Activité*) into a single, consolidated report (CRA n+1). The project processes input PowerPoint (PPTX) files that follow a predefined template available in the `templates` folder and outputs a new summary report in the same format.

## Features

- **Multi-Input Aggregation:**  
  Accepts several CRA PPTX files as input.

- **Content Summarization:**  
  Extracts and summarizes key information from each report.

- **Consolidated Report Generation:**  
  Produces a new CRA (n+1) that combines all the summarized data.

- **Template-Driven Output:**  
  Uses a consistent PPTX template from the `templates` folder to generate the final report.

## Requirements
- Python 3.12
## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/ACRA.git
   cd cra-summary-generator
   ```

2. **Create and activate a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows: venv\Scripts\activate
   ```

3. **Install the required packages:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare Your Input Files:**
   - Place all your CRA PPTX files (the activity reports you want to summarize) in a designated folder (e.g., `pptx_folder/`).

2. **Verify the Template:**
   - Ensure that the template PPTX file is located in the `templates/` folder.

3. **Run the Summarization Script:**
   - Use the provided script `main.py` to generate the consolidated CRA:
   ```bash
   python main.py --input input_reports/ --template templates/template.pptx --output output_summary.pptx
   ```

4. **Review the Output:**
   - The script will create a new PowerPoint file (`output_summary.pptx`) that contains the summary report (CRA n+1).

## Configuration

- The summarization logic and processing can be modified by editing the `generate_summary.py` file.
- Command-line arguments let you specify the input folder, template file, and output file.
- Further customizations (e.g., text extraction, summarization rules) can be implemented based on your project requirements.
