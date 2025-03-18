import os
from pptx import Presentation
import re
import json
from typing import Dict, List, Tuple, Optional

def is_underlined(run):
    """
    Check if a text run is both bold and underlined.
    """
    if not hasattr(run, 'font'):
        return False
    
    # Check if the run is underlined
    return run.font.underline

def get_rgb_color(run):
    """
    Get the RGB color of a text run.
    Returns tuple (R, G, B) or None if color is not accessible.
    """
    if not hasattr(run, 'font') or run.font.color is None:
        return None
    try:
        rgb = run.font.color.rgb
        if rgb is None:
            return None
        return (rgb[0], rgb[1], rgb[2])
    except AttributeError:
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

def extract_table_data_from_slide(slide) -> List[Dict]:
    """
    Extract table data from a slide, focusing on tables with information fields.
    Returns a list of rows with text and formatting information.
    """
    results = []
    
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            
            # Process each row in the table
            for row in table.rows:
                row_data = []
                
                # Process each cell in the row
                for cell in row.cells:
                    if cell.text_frame:
                        cell_data = {
                            "text": "",
                            "paragraphs": []
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
                                is_text_underlined = is_underlined(run)
                                color = get_rgb_color(run)
                                color_type = identify_color_type(color)
                                
                                para_data["text"] += run_text
                                para_data["runs"].append({
                                    "text": run_text,
                                    "bold_and_underlined": is_text_underlined,
                                    "color": color,
                                    "color_type": color_type
                                })
                            
                            cell_data["text"] += para_data["text"] + "\n"
                            cell_data["paragraphs"].append(para_data)
                            
                        cell_data["text"] = cell_data["text"].strip()
                        row_data.append(cell_data)
                    else:
                        row_data.append({"text": "", "paragraphs": []})
                
                results.append(row_data)
    
    return results

def extract_projects_from_table_data(table_data: List[Dict]) -> Dict[str, Dict]:
    """
    Extract project information from processed table data.
    Returns a dictionary with project names as keys and their information as values.
    """
    projects = {}
    current_project = None
    current_info = []
    
    # Process each row in the table data
    for row in table_data:
        for cell in row:
            # Process paragraphs in each cell
            for paragraph in cell.get("paragraphs", []):
                # Process runs in each paragraph
                for i, run in enumerate(paragraph.get("runs", [])):
                    if run["bold_and_underlined"]:
                        # If we already have a project, save its information
                        if current_project:
                            projects[current_project] = {
                                "information": "".join(current_info),
                                "alerts": {
                                    "advancements": projects.get(current_project, {}).get("alerts", {}).get("advancements", []),
                                    "small_alerts": projects.get(current_project, {}).get("alerts", {}).get("small_alerts", []),
                                    "critical_alerts": projects.get(current_project, {}).get("alerts", {}).get("critical_alerts", [])
                                }
                            }
                        
                        # Start a new project
                        current_project = run["text"].strip()
                        current_info = []
                    elif current_project:
                        # Add to the current project's information
                        current_info.append(run["text"])
                        
                        # Track alerts by color
                        if run["color_type"] == "advancement":
                            projects.setdefault(current_project, {}).setdefault("alerts", {}).setdefault("advancements", []).append(run["text"])
                        elif run["color_type"] == "small_alert":
                            projects.setdefault(current_project, {}).setdefault("alerts", {}).setdefault("small_alerts", []).append(run["text"])
                        elif run["color_type"] == "critical_alert":
                            projects.setdefault(current_project, {}).setdefault("alerts", {}).setdefault("critical_alerts", []).append(run["text"])
    
    # Don't forget the last project
    if current_project and current_info:
        projects[current_project] = {
            "information": "".join(current_info),
            "alerts": {
                "advancements": projects.get(current_project, {}).get("alerts", {}).get("advancements", []),
                "small_alerts": projects.get(current_project, {}).get("alerts", {}).get("small_alerts", []),
                "critical_alerts": projects.get(current_project, {}).get("alerts", {}).get("critical_alerts", [])
            }
        }
    
    return projects

def extract_projects_from_presentation(file_path: str) -> Dict[str, Dict]:
    """
    Extract project information from a PowerPoint presentation.
    Focuses on the first slide and tables with information fields.
    """
    try:
        prs = Presentation(file_path)
        
        # Process only the first slide as specified
        if len(prs.slides) > 0:
            slide = prs.slides[0]
            table_data = extract_table_data_from_slide(slide)
            projects = extract_projects_from_table_data(table_data)
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