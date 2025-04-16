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
      - Column 1: Project and subproject names with indentation by level
      - Column 2: Project information with colored text
          * Black: Common information
          * Green: Advancements
          * Orange: Small alerts
          * Red: Critical alerts
      - Column 3: Upcoming events by service
      
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
    
    # Fonction récursive pour ajouter les projets à tous les niveaux
    def add_project_level(projects, level=0):
        nonlocal current_row, table
        
        for project_name, content in projects.items():
            # Déterminer si c'est un niveau terminal (avec des données) ou un container
            is_terminal = "information" in content
            
            # If we need more rows in the table, add them
            while current_row >= len(table.rows):
                add_row(table)
            
            # Calculate indentation based on level
            indent = "  " * level
            
            # Set project name in column 0 with proper indentation
            cell = table.cell(current_row, 0)
            cell.text = f"{indent}{project_name}"
            
            # Apply formatting based on level
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    if level == 0:
                        # Top level projects in bold
                        run.font.bold = True
                    elif level > 1:
                        # Deeper level projects in italic
                        run.font.italic = True
            
            if is_terminal:
                # Process project information for column 1
                info_cell = table.cell(current_row, 1)
                info_cell.text = ""
                
                # Add information with formatted text
                tf = info_cell.text_frame
                
                # Add summary/information text (black)
                if "information" in content and content["information"]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = content["information"]
                
                # Add advancements (green)
                if "advancements" in content and content["advancements"]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(content["advancements"])
                    run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                
                # Add small alerts (orange)
                if "small" in content and content["small"]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(content["small"])
                    run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
                
                # Add critical alerts (red)
                if "critical" in content and content["critical"]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = "\n".join(content["critical"])
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Red
                
                table.rows[current_row].height = Pt(12)
                current_row += 1
            else:
                # Empty cells for columns 1 and 2
                info_cell = table.cell(current_row, 1)
                info_cell.text = ""
                
                events_cell = table.cell(current_row, 2)
                events_cell.text = ""
                
                current_row += 1
                
                # Process next level recursively
                add_project_level(content, level + 1)
    
    # Process all projects recursively
    add_project_level(project_data)
    
    # Ajouter une section pour les événements à venir par service
    if upcoming_events:
        # Si nous sommes à la fin de la table, ajouter une ligne pour le titre
        while current_row >= len(table.rows):
            add_row(table)
        
        # Ajouter les événements par service
        for service_name, events in upcoming_events.items():
            if events:
                while current_row >= len(table.rows):
                    add_row(table)
                
                # Ajouter le nom du service
                service_cell = table.cell(current_row, 0)
                service_cell.text = service_name
                for paragraph in service_cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                
                # Ajouter les événements
                events_cell = table.cell(current_row, 1)
                events_cell.text = ""
                
                tf = events_cell.text_frame
                for event in events:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = event
                
                # # Fusionner les colonnes restantes
                if len(table.columns) >= 3:
                    events_cell.merge(table.cell(current_row, 2))
                
                current_row += 1
    
    # Save the presentation
    prs.save(output_path)
    print(f"Updated table with project data and saved to {output_path}")
    return output_path

