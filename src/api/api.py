import os,sys
import shutil
import uvicorn
from pptx import presentation
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse  
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from core import summarize_ppt, get_slide_structure, get_slide_structure_wcolor, delete_all_pptx_files

# starting Fast API 
app = FastAPI() 
load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")

# curl -X POST "http://localhost:5050/acra/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@CRA_SERVICE_CYBER.pptx"
@app.get("/acra/{folder_name}")
async def summarize(folder_name: str):
    """
    Summarizes the content of PowerPoint files in a folder and updates a template PowerPoint file with the summary.
    The PowerPoint will be structured with a hierarchical format:
      - Main projects as headers
      - Subprojects under each main project
      - Information, alerts for each subproject
      - Events listed by service at the bottom

    Args:
        folder_name (str): The name of the folder containing PowerPoint files to analyze.

    Returns:
        dict: A dictionary containing the download URL of the updated PowerPoint file.

    Raises:
        HTTPException: If there's an error processing the PowerPoint files.
    """
    try:
        return summarize_ppt(folder_name)
    except Exception as e:
        # Log the exception for debugging
        print(f"Error in summarize_folder: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Summarize error: {str(e)}")

# Testing the function with : 
#  curl -X GET "http://localhost:5050/get_slide_structure/CRA_SERVICE_CYBER.pptx"
@app.get("/get_slide_structure/{foldername}")
async def get_structure(foldername: str):
    """
    Analyse tous les fichiers PPTX présents dans le dossier spécifié et fusionne leurs données
    en une structure unique de projets, avec les événements à venir regroupés par service.

    Args:
        foldername (str): Nom du dossier contenant les fichiers PPTX

    Returns:
        dict: Un dictionnaire contenant les projets fusionnés et les événements à venir par service.

    Raises:
        HTTPException: Si aucun fichier PPTX n'est trouvé.
    """
    try:
        return get_slide_structure(foldername)
    except Exception as e:
        print(f"Error in slide_structure : {str(e)}")
        raise HTTPException(status_code=500, detail=f"Slide structure error: {str(e)}")


#  curl -X GET "http://localhost:5050/get_slide_structure_wcolor/CRA_SERVICE_CYBER.pptx"
@app.get("/get_slide_structure_wcolor/{filename}")
async def structure_wcolor(filename: str):
    """
    Endpoint to get the structure of a slide presentation with colors detection.

    Args:
        filename (str): The name of the file to analyze.

    Returns:
        dict: A dictionary containing the filename and the number of slides.

    Raises:
        HTTPException: If the file does not exist, a 404 error is raised.
    """
    try:
        return get_slide_structure_wcolor(filename)
    except Exception as e:
        print(f"Error in slide_structure_wcolor : {str(e)}")
        raise HTTPException(status_code=500, detail=f"Slide structure wcolor error: {str(e)}")

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

@app.delete("/delete_all_pptx_files/{foldername}")
async def delete_files(foldername:str):
    """
    Supprime tous les fichiers du dossier pptx_folder.

    Returns:
        dict: Un message confirmant la suppression des fichiers.
    
    Raises:
        HTTPException: Si le dossier n'existe pas.
    """
    try:
        return delete_all_pptx_files(foldername)
    except Exception as e:
        print(f"Erreur lors de la suppression : {str(e)}")
        raise HTTPException(status_code=500, detail=f"Deletion error : {str(e)}")


def run():
    uvicorn.run(app, host="0.0.0.0", port=5050)


if __name__ == "__main__":
    summarize_ppt("669fa53b-649a-4023-8066-4cd86670e88b")