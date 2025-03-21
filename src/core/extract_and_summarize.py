import os,sys
import re
from pptx import Presentation

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from config import summarize_model
from analist import extract_projects_from_presentation
from OLLibrary.utils.text_service import remove_tags_no_keep

def extract_if_from_pptx(file_path, aggregated_sections, title_pattern, expected_titles):
    """
    Process a single PPTX file and extract the information fields (IF) from the target table.
    The table is expected to have its header row (titles) and the next row (IF) either as separate cells
    or merged into one cell (with newline-separated values).
    """
    prs = Presentation(file_path)
    for slide in prs.slides:
        # Identify the slide by its title matching the pattern.
        slide_title = None
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                if title_pattern.match(shape.text.strip()):
                    slide_title = shape.text.strip()
                    break
        if not slide_title:
            continue

        # Search for the table (GraphicFrame with a table) with the expected headers.
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                if len(table.rows) < 2:
                    continue

                # Extract header texts.
                header_cells = table.rows
                row_title = ""
                for row_index, row in enumerate(header_cells):
                    row_text = [cell.text for cell in row.cells]
                    if row_index % 2 == 1:
                        aggregated_sections[row_title].append(row_text[0])
                    else:
                        row_title = row_text[0]

def aggregate_pptx_files(pptx_folder):
    """
    Loop through all PPTX files in the folder, extract IF texts for each file,
    and return a dictionary aggregating texts per section.
    """
    aggregated_sections = {}
    for filename in os.listdir(pptx_folder):
        if filename.lower().endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            aggregated_sections = extract_projects_from_presentation(file_path)
    # Combine the IF texts for each section into one string.
    for key in aggregated_sections:
        aggregated_sections[key] = "\n".join(aggregated_sections[key])
    return aggregated_sections

def extract_common_and_upcoming_info(project_data):
    """
    Extract common information and upcoming work information from project data.
    
    Parameters:
      project_data (dict): Project data dictionary extracted from presentations.
    
    Returns:
      tuple: (common_info, upcoming_info) where both are strings
    """
    common_info = []
    upcoming_info = ""
    
    # Extract common information from all projects
    for project_name, project_info in project_data.items():
        if "information" in project_info:
            info_text = project_info["information"]
            
            # Check if the information contains details about upcoming week
            upcoming_week_match = re.search(r"Evénements de la semaine à venir(.*?)$", info_text, re.DOTALL)
            if upcoming_week_match:
                # Split the information: before match goes to common_info, the match itself goes to upcoming_info
                common_part = info_text[:upcoming_week_match.start()]
                upcoming_part = upcoming_week_match.group(1).strip()
                
                if common_part:
                    common_info.append(f"{project_name}: {common_part}")
                if upcoming_part:
                    upcoming_info += f"{project_name}: {upcoming_part}\n"
            else:
                common_info.append(f"{project_name}: {info_text}")
        
        # Also include alerts in common information
        if "alerts" in project_info:
            alerts = project_info["alerts"]
            alert_text = []
            
            if alerts.get("advancements"):
                alert_text.append("Avancements: " + ", ".join(alerts["advancements"]))
            
            if alerts.get("small_alerts"):
                alert_text.append("Alertes mineures: " + ", ".join(alerts["small_alerts"]))
            
            if alerts.get("critical_alerts"):
                alert_text.append("Alertes critiques: " + ", ".join(alerts["critical_alerts"]))
            
            if alert_text:
                common_info.append(f"{project_name} (Alertes): {' | '.join(alert_text)}")
    
    return "\n\n".join(common_info), upcoming_info if upcoming_info else "Aucun événement particulier prévu pour la semaine à venir."

def summarize_sections(aggregated_sections):
    """
    Build the prompt from the aggregated sections and call the ChatOpenAI model to obtain a summary.
    """
    prompt = (
        "Here is aggregated information from multiple PowerPoint files. Your task is to create a structured summary in French, dividing the content by each relevant section. For each section, provide a concise overview that captures the key points and main ideas. If you encounter sections with minimal or irrelevant updates, you may briefly mention them or skip them entirely."
        "Key Requirements:"
        "   - Language: The summary must be written entirely in French."
        "   - Structure: Separate the content by section, ensuring each section has a clear heading or title."
        "   - Conciseness: Deliver a concise yet informative summary, focusing on essential points."
        "   - Relevancy: If any section does not contain substantial information (e.g., \"no big updates\"), you may either omit it or note it briefly."
        "   - Accuracy: Maintain the integrity of the original information; do not add unsupported details."
        "Output Format:"
        "   - One consolidated summary in French, broken down by each section from the source material."
        "   - Use headings or bullet points to keep the sections clear and organized."
    )
    for section, text in aggregated_sections.items():
        prompt += f"=== {section} ===\n{text}\n\n"

    summary = summarize_model.invoke(prompt)
    return summary.content

def aggregate_and_summarize(pptx_folder):
    """
    Main function to aggregate the IF texts from all PPTX files in the folder and obtain a summarized result.
    Returns a tuple of (common_info, upcoming_info) for updating the PowerPoint template.
    """
    project_data = {}
    
    # Get all PPTX files in the folder
    for filename in os.listdir(pptx_folder):
        if filename.lower().endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            # Extract project data from the presentation
            file_project_data = extract_projects_from_presentation(file_path)
            # Merge with existing project data
            for project_name, project_info in file_project_data.items():
                project_data[project_name] = project_info
    
    # Extract common information and upcoming work information
    common_info, upcoming_info = extract_common_and_upcoming_info(project_data)
    
    # Summarize common information if it's too long
    if len(common_info) > 2000:  # Arbitrary threshold, adjust as needed
        prompt = (
            "Voici des informations agrégées à partir de plusieurs fichiers PowerPoint. "
            "Votre tâche est de créer un résumé structuré en français qui capture les points clés "
            "et les idées principales de manière concise et informative.\n\n"
            f"{common_info}"
        )
        common_info = summarize_model.invoke(prompt).content
        common_info = remove_tags_no_keep(common_info, "<think>", "</think>")
    
    # Summarize upcoming information if it's too long
    if len(upcoming_info) > 1000:  # Arbitrary threshold, adjust as needed
        prompt = (
            "Voici des informations concernant les événements de la semaine à venir. "
            "Veuillez résumer ces informations de manière concise en français, en conservant "
            "les points essentiels concernant les prochaines étapes et événements.\n\n"
            f"{upcoming_info}"
        )
        upcoming_info = summarize_model.invoke(prompt).content
        upcoming_info = remove_tags_no_keep(upcoming_info, "<think>", "</think>")
    
    return {"common_info": common_info, "upcoming_info": upcoming_info}

if __name__ == "__main__":
    folder = "pptx_folder"  # Update with your actual folder path
    aggregate_and_summarize(folder)
