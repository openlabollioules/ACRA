import os,sys
import shutil
import uvicorn
from pptx import presentation
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse  

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from analist import analyze_presentation , analyze_presentation_with_colors
from services import update_table_cell
from core import aggregate_and_summarize
# starting Fast API 
app = FastAPI() 

UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER")

# curl -X POST "http://localhost:5050/acra/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@CRA_SERVICE_CYBER.pptx"
@app.post("/acra/")
async def summarize_ppt(file: UploadFile = File(...)):
    """
    Summarizes the content of an uploaded PowerPoint file and updates a template PowerPoint file with the summary.

    Args:
        file (UploadFile): The uploaded PowerPoint file to be summarized.

    Returns:
        dict: A dictionary containing the download URL of the updated PowerPoint file.

    Raises:
        HTTPException: If there's an error processing the PowerPoint file.
    """
    try:
        # Ensure the upload directory exists
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        
        # Save the uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Generate the summary from the PowerPoint files in the upload folder
        summary_text = aggregate_and_summarize(UPLOAD_FOLDER)
        
        if not summary_text:
            raise HTTPException(status_code=400, detail="Le PowerPoint est vide ou n'a pas de texte extractible.")
        
        # Update the template with the summary text
        output_filename = f"{OUTPUT_FOLDER}/updated_presentation.pptx"
        summarized_file_path = update_table_cell(
            pptx_path=os.getenv("TEMPLATE_FILE"),  # Template file 
            slide_index=0,  # first slide
            table_shape_index=1,  # index of the table
            row=1,  # Write inside the row 1 of the table (title area in row: 0,2,4)
            col=0, 
            new_text=summary_text, 
            output_path=output_filename
        )

        # Clean up the upload folder - safely remove files first
        for filename in os.listdir(UPLOAD_FOLDER):
            file_to_remove = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(file_to_remove):
                os.remove(file_to_remove)
        
        return {"download_url": f"/download/{summarized_file_path}"}
    
    except Exception as e:
        # Log the exception for debugging
        print(f"Error in summarize_ppt: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Une erreur s'est produite: {str(e)}")

# Testing the function with : 
#  curl -X GET "http://localhost:5050/get_slide_structure/CRA_SERVICE_CYBER.pptx"
@app.get("/get_slide_structure/{filename}")
async def get_slide_structure(filename: str):
    """
    Endpoint to get the structure of a slide presentation.

    Args:
        filename (str): The name of the file to analyze.

    Returns:
        dict: A dictionary containing the filename and the number of slides.

    Raises:
        HTTPException: If the file does not exist, a 404 error is raised.
    """
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    slides_data = analyze_presentation(file_path)
    return {"filename": filename, "slide data": slides_data}

#  curl -X GET "http://localhost:5050/get_slide_structure_wcolor/CRA_SERVICE_CYBER.pptx"
@app.get("/get_slide_structure_wcolor/{filename}")
async def get_slide_structure_wcolor(filename: str):
    """
    Endpoint to get the structure of a slide presentation with colors detection.

    Args:
        filename (str): The name of the file to analyze.

    Returns:
        dict: A dictionary containing the filename and the number of slides.

    Raises:
        HTTPException: If the file does not exist, a 404 error is raised.
    """
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    slides_data = analyze_presentation_with_colors(file_path)
    return {"filename": filename, "slide data": slides_data}

# Testing the function with : 
#  curl -OJ http://localhost:5050/download/TEST_FILE.pptx
@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Download a file from the server.
    This endpoint allows clients to download a file from the server by specifying the filename in the URL path.
    Args:
        filename (str): The name of the file to be downloaded.
    Returns:
        FileResponse: A response containing the file to be downloaded.
    Raises:
        HTTPException: If the file does not exist, a 404 status code is returned with a detail message "File Not found.".
    """
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File Not found.")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


def run():
    uvicorn.run(app, host="0.0.0.0", port=5050)