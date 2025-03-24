import os,sys
import re
from pptx import Presentation

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from config import summarize_model
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

def aggregate_pptx_files(pptx_folder, title_pattern, expected_titles):
    """
    Loop through all PPTX files in the folder, extract IF texts for each file,
    and return a dictionary aggregating texts per section.
    """
    aggregated_sections = {title: [] for title in expected_titles}
    for filename in os.listdir(pptx_folder):
        if filename.lower().endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            extract_if_from_pptx(file_path, aggregated_sections, title_pattern, expected_titles)
    # Combine the IF texts for each section into one string.
    for key in aggregated_sections:
        aggregated_sections[key] = "\n".join(aggregated_sections[key])
    return aggregated_sections

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
        "   - Relevancy: If any section does not contain substantial information (e.g., “no big updates”), you may either omit it or note it briefly."
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
    """
    title_pattern = "whessss"
    expected_titles = [
        "Activités de la semaine",
        "Alertes et Points durs",
        "Evénements de la semaine à venir"
    ]
    aggregated_sections = aggregate_pptx_files(pptx_folder, title_pattern, expected_titles)
    summary = remove_tags_no_keep(summarize_sections(aggregated_sections), "<think>", "</think>")
    return summary

# if __name__ == "__main__":
#     folder = "pptx_folder"  
#     aggregate_and_summarize(folder)
