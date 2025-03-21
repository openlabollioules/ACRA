from pptx import Presentation
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
    
    # Save the updated presentation
    prs.save(output_path)
    print(f"Updated multiple cells and saved to {output_path}")
    return output_path

