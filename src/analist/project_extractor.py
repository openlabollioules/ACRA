import os
from pptx import Presentation
import re
import json
from typing import Dict, List, Tuple, Optional

def is_underlined(run):
    """
    Check if a text run is underlined.
    """
    # Check if run has font attribute directly
    if hasattr(run, 'font') and hasattr(run.font, 'underline'):
        return run.font.underline
    
    # Check if run has parent with font attribute
    if hasattr(run, '_parent') and hasattr(run._parent, 'font') and hasattr(run._parent.font, 'underline'):
        return run._parent.font.underline
    
    return False

def is_bold(run):
    """
    Check if a text run is bold.
    """
    # Check if run has font attribute directly
    if hasattr(run, 'font'):
        # Check for bold in font name or bold attribute
        if (hasattr(run.font, 'name') and run.font.name and "bold" in run.font.name.lower()) or \
           (hasattr(run.font, 'bold') and run.font.bold):
            return True
    
    # Check if run has parent with font attribute
    if hasattr(run, '_parent') and hasattr(run._parent, 'font'):
        # Check for bold in parent font name or bold attribute
        if (hasattr(run._parent.font, 'name') and run._parent.font.name and "bold" in run._parent.font.name.lower()) or \
           (hasattr(run._parent.font, 'bold') and run._parent.font.bold):
            return True
    
    return False

def get_rgb_color(run):
    """
    Get the RGB color of a text run.
    Returns tuple (R, G, B) or None if color is not accessible.
    """
    # Try to get color from run's font
    if hasattr(run, 'font') and hasattr(run.font, 'color') and run.font.color is not None:
        if hasattr(run.font.color, 'rgb') and run.font.color.rgb is not None:
            return tuple(run.font.color.rgb)
    
    # Try to get color from run's parent font
    if hasattr(run, '_parent') and hasattr(run._parent, 'font') and \
       hasattr(run._parent.font, 'color') and run._parent.font.color is not None:
        if hasattr(run._parent.font.color, 'rgb') and run._parent.font.color.rgb is not None:
            return tuple(run._parent.font.color.rgb)
    
    return None

def identify_color_type(color_tuple: Tuple[int, int, int]) -> str:
    """
    Identify color type based on RGB values.
    - Green: big advancement
    - Orange: small alert
    - Red: critical alert
    """
    if color_tuple is None:
        return "normal"
    
    r, g, b = color_tuple
    
    # Simple heuristic for color identification
    if g > max(r, b) + 50:  # Green is dominant
        return "advancement"
    elif r > g + 50 and g > b + 50:  # Orange-ish
        return "small_alert"
    elif r > max(g, b) + 50:  # Red is dominant
        return "critical_alert"
    else:
        return "normal"

def extract_title_from_slide(slide) -> str:
    """
    Extract title from a text field in the slide.
    Returns the title text.
    """
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            # Assuming the first text field with content is the title
            if shape.text_frame.text.strip():
                return shape.text_frame.text.strip()
    return "Untitled"

def extract_table_data_from_slide(slide) -> List[Dict]:
    """
    Extract table data from a slide, focusing on tables with 3 columns:
    1. Project name
    2. Project information
    3. Upcoming events
    Returns a list of rows with text and formatting information.
    """
    results = []
    
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            
            # Verify that we have the expected table structure - at least 3 columns
            if len(table.columns) < 3:
                print(f"Warning: Table does not have 3 columns (found {len(table.columns)}). Skipping.")
                continue
            
            # Process each row in the table
            for row in table.rows:
                # Skip header row if it exists (optional)
                # You can uncomment this if your table has a header row to skip
                # if row_idx == 0:
                #     continue
                
                row_data = []
                
                # We only care about the 3 columns we expect - but avoid slicing
                for col_idx, cell in enumerate(row.cells):
                    # Only process the first 3 columns
                    if col_idx >= 3:
                        break
                        
                    if cell.text_frame:
                        cell_data = {
                            "text": "",
                            "paragraphs": [],
                            "column_index": col_idx  # Track which column this is
                        }
                        
                        # Process each paragraph in the cell
                        for paragraph in cell.text_frame.paragraphs:
                            para_data = {
                                "text": "",
                                "runs": []
                            }
                            
                            # Process each run in the paragraph
                            for run in paragraph.runs:
                                run_text = run.text
                                color = get_rgb_color(run)
                                color_type = identify_color_type(color)
                                
                                para_data["text"] += run_text
                                para_data["runs"].append({
                                    "text": run_text,
                                    "color": color,
                                    "color_type": color_type
                                })
                            
                            cell_data["text"] += para_data["text"] + "\n"
                            cell_data["paragraphs"].append(para_data)
                            
                        cell_data["text"] = cell_data["text"].strip()
                        row_data.append(cell_data)
                    else:
                        row_data.append({"text": "", "paragraphs": [], "column_index": col_idx})
                
                # Only add rows that have some content
                if any(cell.get("text", "").strip() for cell in row_data):
                    results.append(row_data)
    
    return results

def extract_projects_from_table_data(table_data: List[Dict], title: str) -> Dict[str, Dict]:
    """
    Extract project information from processed table data.
    Assuming the table has 3 columns:
    1. Project name (column_index=0)
    2. Project information with colored text for alerts (column_index=1)
    3. Upcoming events (column_index=2)
    Returns a dictionary with project names as keys and their information as values.
    """
    projects = {}
    upcoming_events = []
    
    # Initialize the title as the main key - use "activities" as the standard key
    projects["activities"] = {}
    
    # Process each row in the table data
    for row in table_data:
        # Get cells by column index
        project_name_cell = next((cell for cell in row if cell.get("column_index") == 0), {})
        project_info_cell = next((cell for cell in row if cell.get("column_index") == 1), {})
        events_cell = next((cell for cell in row if cell.get("column_index") == 2), {})
        
        # Extract project name from column 0
        project_name = project_name_cell.get("text", "").strip()
        if project_name:
            # Initialize project data if this is a new project
            if project_name not in projects["activities"]:
                projects["activities"][project_name] = {
                    "information": "",
                    "alerts": {
                        "advancements": [],
                        "small_alerts": [],
                        "critical_alerts": []
                    }
                }
            
            # Process project information from column 1
            for paragraph in project_info_cell.get("paragraphs", []):
                projects["activities"][project_name]["information"] += paragraph.get("text", "") + "\n"
                
                # Process runs to extract colored alerts
                for run in paragraph.get("runs", []):
                    if run["color_type"] == "advancement":
                        projects["activities"][project_name]["alerts"]["advancements"].append(run["text"])
                    elif run["color_type"] == "small_alert":
                        projects["activities"][project_name]["alerts"]["small_alerts"].append(run["text"])
                    elif run["color_type"] == "critical_alert":
                        projects["activities"][project_name]["alerts"]["critical_alerts"].append(run["text"])
            
            # Clean up information text
            projects["activities"][project_name]["information"] = projects["activities"][project_name]["information"].strip()
        
        # Process the upcoming events from column 2 (collect from all rows with data)
        if events_cell.get("text", "").strip() and not events_cell.get("text") in upcoming_events:
            for paragraph in events_cell.get("paragraphs", []):
                paragraph_text = paragraph.get("text", "").strip()
                if paragraph_text and paragraph_text not in upcoming_events:
                    upcoming_events.append(paragraph_text)
    
    # Add upcoming events to the projects dictionary
    if upcoming_events:
        projects["upcoming_events"] = "\n".join(upcoming_events).strip()
    
    # Store the title as metadata
    projects["metadata"] = {
        "title": title
    }
    
    return projects

def extract_projects_from_presentation(file_path: str) -> Dict[str, Dict]:
    """
    Extract project information from a PowerPoint presentation.
    Focuses on the first slide with a title and a 3-column table.
    """
    try:
        prs = Presentation(file_path)
        
        # Process only the first slide as specified
        if len(prs.slides) > 0:
            slide = prs.slides[0]
            title = extract_title_from_slide(slide)
            table_data = extract_table_data_from_slide(slide)
            projects = extract_projects_from_table_data(table_data, title)
            return projects
        else:
            return {}
    except Exception as e:
        print(f"Error processing presentation: {e}")
        return {}

def format_projects_as_json(projects: Dict[str, Dict], output_file: Optional[str] = None) -> str:
    """
    Format project information as JSON and optionally save to a file.
    """
    json_data = json.dumps(projects, indent=2, ensure_ascii=False)
    
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_data)
    
    return json_data

def extract_and_format_projects(file_path: str, output_file: Optional[str] = None) -> Dict[str, Dict]:
    """
    Main function to extract project information from a PowerPoint presentation
    and format it as JSON.
    """
    projects = extract_projects_from_presentation(file_path)
    
    if output_file:
        format_projects_as_json(projects, output_file)
    
    return projects

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pptx_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        projects = extract_and_format_projects(pptx_file, output_file)
        
        if not output_file:
            print(json.dumps(projects, indent=2, ensure_ascii=False))
    else:
        print("Usage: python project_extractor.py <pptx_file> [output_json_file]") 