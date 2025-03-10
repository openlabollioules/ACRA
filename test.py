from slide_analist import analyze_presentation
from update_pttx_service import update_table_cell

analyze_presentation()

# Assuming the table is in the first slide (index 0) and is the third shape (index 2)
update_table_cell(
    pptx_path="./templates/CRA_LA_GT_TEMPLATE.pptx", 
    slide_index=0, 
    table_shape_index=2, 
    row=1, 
    col=0, 
    new_text="Updated Text will be here", 
    output_path="updated_presentation.pptx"
)

