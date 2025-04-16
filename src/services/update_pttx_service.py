from pptx import Presentation
from pptx.dml.color import RGBColor
from copy import deepcopy
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
import os

def update_table_cell(pptx_path, slide_index, table_shape_index, row, col, new_text, output_path):
    """
    Updates the text of a cell in a table within a PowerPoint file.
    
    Parameters:
      pptx_path (str): Path to the input .pptx file.
      slide_index (int): Index of the slide containing the table.
      table_shape_index (int): Index of the shape that is the table on that slide.
      row (int): Row index of the cell (0-indexed).
      col (int): Column index of the cell (0-indexed).
      new_text (str): New text to insert into the cell.
      output_path (str): Path to save the updated .pptx file.
    """
    os.makedirs(os.getenv("OUTPUT_FOLDER", "OUTPUT"), exist_ok=True)
    # Load the presentation
    prs = Presentation(pptx_path)
    
    # Access the specified slide
    slide = prs.slides[slide_index]
    
    # Access the shape that contains the table
    table_shape = slide.shapes[table_shape_index]
    
    # Check if the shape contains a table
    if not table_shape.has_table:
        raise ValueError("The specified shape does not contain a table.")
    
    # Access the table
    table = table_shape.table
    
    # Update the text in the specified cell
    table.cell(row, col).text = new_text
    
    # Save the updated presentation
    prs.save(output_path)
    print(f"Updated cell ({row}, {col}) with text: '{new_text}' and saved to {output_path}")
    return output_path

def update_table_multiple_cells(pptx_path, slide_index, table_shape_index, updates, output_path):
    """
    Updates multiple cells in a table within a PowerPoint file.
    
    Parameters:
      pptx_path (str): Path to the input .pptx file.
      slide_index (int): Index of the slide containing the table.
      table_shape_index (int): Index of the shape that is the table on that slide.
      updates (list): List of dictionaries with row, col, and text keys.
      output_path (str): Path to save the updated .pptx file.
    
    Example of updates parameter:
    [
        {'row': 1, 'col': 0, 'text': 'Common information text'},
        {'row': 3, 'col': 0, 'text': 'Upcoming work information'}
    ]
    """
    os.makedirs(os.getenv("OUTPUT_FOLDER", "OUTPUT"), exist_ok=True)
    # Load the presentation
    prs = Presentation(pptx_path)
    
    # Access the specified slide
    slide = prs.slides[slide_index]
    
    # Access the shape that contains the table
    table_shape = slide.shapes[table_shape_index]
    
    # Check if the shape contains a table
    if not table_shape.has_table:
        raise ValueError("The specified shape does not contain a table.")
    
    # Access the table
    table = table_shape.table
    
    # Update each cell as specified
    for update in updates:
        row = update.get('row')
        col = update.get('col')
        text = update.get('text')
        
        if row is not None and col is not None and text is not None:
            table.cell(row, col).text = text
            print(f"Updated cell ({row}, {col}) with text")
    
    # Save the presentation
    prs.save(output_path)
    print(f"Updated multiple cells and saved to {output_path}")
    return output_path

def add_row(table):
    """
    Copie la dernière ligne du tableau et l'ajoute à la fin.
    """
    # Copie de la dernière ligne
    new_row = deepcopy(table._tbl.tr_lst[-1])
    # Ajoute la nouvelle ligne au tableau
    table._tbl.append(new_row)

def merge_vertical(first_cell, last_cell):
    """
    Fusionne verticalement une liste de cellules.
    """
    # Pour la première cellule
    # for cell in cells:
    first_cell.merge(last_cell)

def update_table_with_project_data(pptx_path, slide_index, table_shape_index, project_data, output_path, upcoming_events=None):
    """
    Updates a table in a PowerPoint slide with project information using the new nested JSON format.
    Supports colored text for different types of alerts and multi-level project hierarchies.
    
    Parameters:
      pptx_path (str): Path to the input .pptx file.
      slide_index (int): Index of the slide containing the table.
      table_shape_index (int): Index of the shape that is the table on that slide.
      project_data (dict): Project data in the nested format with multi-level hierarchy.
                          Should be directly the content of the "projects" field.
      output_path (str): Path to save the updated .pptx file.
      upcoming_events (dict, optional): Dictionary of upcoming events by service.
    
    The table is organized as follows:
      - Column 1: Top-level project names only
      - Column 2: Project information with subprojects in bold, subsubprojects underlined
          * Black: Common information
          * Green: Advancements
          * Orange: Small alerts
          * Red: Critical alerts
      - Column 3: Upcoming events by service (service names in bold)
      
    Returns:
      str: Path to the saved output file
    """
    os.makedirs(os.getenv("OUTPUT_FOLDER", "OUTPUT"), exist_ok=True)
    
    # Load the presentation
    prs = Presentation(pptx_path)
    
    # Access the specified slide
    slide = prs.slides[slide_index]
    
    # Access the shape that contains the table
    while not slide.shapes[table_shape_index].has_table:
        table_shape_index += 1
    
    # Access the table
    table = slide.shapes[table_shape_index].table
    
    # Start from row 1 (assuming row 0 might be headers)
    current_row = 1
    first_project_row = current_row  # Remember the first row where we start adding projects
    
    # Process each top-level project
    for project_name, project_content in project_data.items():
        # If we need more rows in the table, add them
        while current_row >= len(table.rows):
            add_row(table)
        
        # Set project name in column 1
        cell = table.cell(current_row, 0)
        cell.text = project_name
        
        # Apply bold formatting to top level project names
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER  # Center-align text in first column
            for run in paragraph.runs:
                run.font.bold = True
        
        # Create text frame for column 2 which will contain all project information
        info_cell = table.cell(current_row, 1)
        info_cell.text = ""
        tf = info_cell.text_frame
        tf.clear()
        
        # Add top-level project information if it exists
        if "information" in project_content:
            # Use the first paragraph that already exists in the text frame instead of creating a new one
            if tf.paragraphs:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT  # Left-align text
            run = p.add_run()
            run.font.size = Pt(8)
            run.text = project_content["information"]
            
            # Track if we need to add a paragraph for subsequent content
            has_content = True
            
            # Add top-level project alerts and advancements
            # Add advancements (green)
            if "advancements" in project_content and project_content["advancements"]:
                p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT  # Left-align text
                run = p.add_run()
                run.font.size = Pt(8)
                run.text = "\n".join(project_content["advancements"])
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                has_content = True
            
            # Add small alerts (orange)
            if "small" in project_content and project_content["small"]:
                p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT  # Left-align text
                run = p.add_run()
                run.font.size = Pt(8)
                run.text = "\n".join(project_content["small"])
                run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
                has_content = True
            
            # Add critical alerts (red)
            if "critical" in project_content and project_content["critical"]:
                p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT  # Left-align text
                run = p.add_run()
                run.font.size = Pt(8)
                run.text = "\n".join(project_content["critical"])
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
                has_content = True
        else:
            has_content = False
        
        # Process subprojects recursively
        for subproject_name, subproject_content in project_content.items():
            # Skip non-dictionary fields (already processed)
            if not isinstance(subproject_content, dict) or subproject_name in ["information", "critical", "small", "advancements"]:
                continue
            
            # Add subproject name in bold
            if has_content:
                p = tf.add_paragraph()
            else:
                # Use existing first paragraph if this is the first content
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                has_content = True
            
            p.alignment = PP_ALIGN.LEFT  # Left-align text
            run = p.add_run()
            run.font.size = Pt(8)
            run.text = f"{subproject_name} : "
            run.font.bold = True
            
            # Add subproject information
            if "information" in subproject_content:
                run = p.add_run()
                run.font.size = Pt(8)
                run.text = subproject_content["information"]
                
                # Add subproject alerts and advancements
                # Add advancements (green)
                if "advancements" in subproject_content and subproject_content["advancements"]:
                    run = p.add_run()
                    run.font.size = Pt(8)
                    run.text = "\n".join(subproject_content["advancements"])
                    run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                
                # Add small alerts (orange)
                if "small" in subproject_content and subproject_content["small"]:
                    run = p.add_run()
                    run.font.size = Pt(8)
                    run.text = "\n".join(subproject_content["small"])
                    run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
                
                # Add critical alerts (red)
                if "critical" in subproject_content and subproject_content["critical"]:
                    run = p.add_run()
                    run.font.size = Pt(8)
                    run.text = "\n".join(subproject_content["critical"])
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            
            # Process subsubprojects
            for subsubproject_name, subsubproject_content in subproject_content.items():
                # Skip non-dictionary fields (already processed)
                if not isinstance(subsubproject_content, dict) or subsubproject_name in ["information", "critical", "small", "advancements"]:
                    continue
                
                run = p.add_run()
                run.font.size = Pt(8)
                run.text = f"{subsubproject_name} : "
                run.font.underline = True
                
                # Add subsubproject information
                if "information" in subsubproject_content:
                    run = p.add_run()
                    run.font.size = Pt(8)
                    run.text = subsubproject_content["information"]
                    
                    # Add subsubproject alerts and advancements
                    # Add advancements (green)
                    if "advancements" in subsubproject_content and subsubproject_content["advancements"]:
                        run = p.add_run()
                        run.font.size = Pt(8)
                        run.text = "\n".join(subsubproject_content["advancements"])
                        run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                    
                    # Add small alerts (orange)
                    if "small" in subsubproject_content and subsubproject_content["small"]:
                        run = p.add_run()
                        run.font.size = Pt(8)
                        run.text = "\n".join(subsubproject_content["small"])
                        run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
                    
                    # Add critical alerts (red)
                    if "critical" in subsubproject_content and subsubproject_content["critical"]:
                        run = p.add_run()
                        run.font.size = Pt(8)
                        run.text = "\n".join(subsubproject_content["critical"])
                        run.font.color.rgb = RGBColor(255, 0, 0)  # Red
        
        # Move to the next row
        table.rows[current_row].height = Pt(8)
        current_row += 1
    
    last_project_row = current_row - 1  # Remember the last row of projects
    
    # Handle upcoming events in the third column (if available)
    if upcoming_events and len(table.columns) >= 3:
        print(f"Processing upcoming events for column 3, rows {first_project_row} to {last_project_row}")
        
        # First, prepare content for the merged cell
        events_content = "Upcoming Events\n\n"
        for service_name, events in upcoming_events.items():
            if events:
                events_content += f"{service_name}\n"
                for event in events:
                    events_content += f"• {event}\n"
                events_content += "\n"
        
        # Clear existing content from all cells in column 3
        for row in range(first_project_row, last_project_row + 1):
            table.cell(row, 2).text = ""
        
        # Now perform the merge of all cells in column 3 at once
        try:
            # Only attempt merge if we have multiple cells
            if last_project_row > first_project_row:
                print(f"Attempting to merge all {last_project_row + 1} cells in column 3 at once")
                table.cell(first_project_row, 2).merge(table.cell(last_project_row, 2))
                print("Successfully merged all cells in column 3")
            else:
                print("Only one cell in column 3, no merging needed")
            
            # Now add content to the merged cell
            events_cell = table.cell(first_project_row, 2)
            events_cell.text = ""  # Clear any existing text
            tf = events_cell.text_frame
            tf.clear()
            
            # Add each service and its events
            first_paragraph = True
            for service_name, events in upcoming_events.items():
                if events:
                    # Add service name
                    p = tf.add_paragraph() if not first_paragraph else tf.paragraphs[0]
                    first_paragraph = False
                    p.alignment = PP_ALIGN.LEFT  # Left-align text
                    p.space_before = Pt(6)  # Add some space before each service
                    run = p.add_run()
                    run.text = service_name
                    run.font.bold = True
                    run.font.size = Pt(8)  # Set font size to 8pt
                    
                    # Add events for this service
                    for event in events:
                        p = tf.add_paragraph()
                        p.alignment = PP_ALIGN.LEFT  # Left-align text
                        p.level = 1  # Indent the events under the service name
                        run = p.add_run()
                        run.text = "• " + event  # Add a bullet point for each event
                        run.font.size = Pt(8)  # Set font size to 8pt
            
        except Exception as e:
            print(f"Error during cell merging in column 3: {str(e)}")
            # Fallback: just put content in the first cell
            events_cell = table.cell(first_project_row, 2)
            events_cell.text = events_content
            # Set left alignment for fallback text
            for paragraph in events_cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
    
    # Save the presentation
    prs.save(output_path)
    print(f"Updated table with project data and saved to {output_path}")
    return output_path

