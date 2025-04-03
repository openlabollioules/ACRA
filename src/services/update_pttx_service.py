from pptx import Presentation
from pptx.dml.color import RGBColor
from copy import deepcopy
from pptx.util import Pt
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
    os.makedirs(os.getenv("OUTPUT_FOLDER"), exist_ok=True)
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
    os.makedirs(os.getenv("OUTPUT_FOLDER"), exist_ok=True)
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
    Puis, dans la nouvelle ligne, modifie la cellule de la 3ème colonne
    pour qu'elle continue la fusion avec les cellules déjà fusionnées.
    """
    # Copie de la dernière ligne
    new_row = deepcopy(table._tbl.tr_lst[-1])
    # Ajoute la nouvelle ligne au tableau
    table._tbl.append(new_row)

def merge_vertical(first_cell, cells):
    """
    Fusionne verticalement une liste de cellules.
    """
    # Pour la première cellule
    for cell in cells:
        first_cell.merge(cell)

def update_table_with_project_data(pptx_path, slide_index, table_shape_index, project_data, output_path):
    """
    Updates a table in a PowerPoint slide with project information using the new JSON format.
    Supports colored text for different types of alerts.
    
    Parameters:
      pptx_path (str): Path to the input .pptx file.
      slide_index (int): Index of the slide containing the table.
      table_shape_index (int): Index of the shape that is the table on that slide.
      project_data (dict): Project data in the new JSON format with activities and upcoming_events.
      output_path (str): Path to save the updated .pptx file.
    
    The table is organized as follows:
      - Column 1: Project names (each project in a separate row)
      - Column 2: Project information with colored text
          * Black: Common information
          * Green: Advancements
          * Orange: Small alerts
          * Red: Critical alerts
      - Column 3: Upcoming events (all in one row)
      
    Returns:
      str: Path to the saved output file
    """
    os.makedirs(os.getenv("OUTPUT_FOLDER"), exist_ok=True)
    
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
    current_row, row_start = 1, 1

    cell_to_merge = []
    
    # Process each project and add to the table
    if "activities" in project_data:
        for project_name, project_info in project_data["activities"].items():
            # If we need more rows in the table, add them
            while current_row >= len(table.rows):
                add_row(table)
                cell_to_merge += [table.rows[current_row].cells[2]]
            
            # Set project name in column 0
            cell = table.cell(current_row, 0)
            cell.text = project_name
            # Process project information for column 1
            info_cell = table.cell(current_row, 1)
            # Clear existing text
            info_cell.text = ""
            # Add information with formatted text
            tf = info_cell.text_frame

            # Add summary text (black)
            if "summary" in project_info:
                p = tf.add_paragraph()
                run = p.add_run()
                run.text = project_info["summary"]
                run.font.size = Pt(11)
            
            # Add alerts with appropriate colors
            if "alerts" in project_info:
                alerts = project_info["alerts"]
                
                # Add advancements (green)
                if alerts.get("advancements") and len(alerts["advancements"]) > 0:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(alerts["advancements"])
                    run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                
                # Add small alerts (orange)
                if alerts.get("small_alerts") and len(alerts["small_alerts"]) > 0:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(alerts["small_alerts"])
                    run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
                
                # Add critical alerts (red)
                if alerts.get("critical_alerts") and len(alerts["critical_alerts"]) > 0:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(alerts["critical_alerts"])
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            table.rows[current_row].height = Pt(11)
            current_row += 1
    row_end = current_row - 1
    if row_end > row_start:
        table.cell(row_start, 2).merge(table.cell(row_end, 2))
    # Add upcoming events in column 2
    if "upcoming_events" in project_data:
        # Add events text to first row, column 2
        events_cell = table.cell(1, 2)
        events_cell.text = ""
        tf = events_cell.text_frame
        p = tf.add_paragraph()
        run = p.add_run()
        
        # Add each event category 
        for category, event_text in project_data["upcoming_events"].items():
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"{category}: {event_text}"
            p.level = 1  # Add a bit of indentation

    # Set the font size to 11 for all runs in the table
    for row in table.rows:
        for cell in row.cells:
            if cell.text_frame:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
    # Save the presentation
    prs.save(output_path)
    print(f"Updated table with project data and saved to {output_path}")
    return output_path

