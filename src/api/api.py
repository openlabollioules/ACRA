import os,sys
import shutil
import uvicorn
import uuid
import logging
from pptx import presentation
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse  
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from core import summarize_ppt, get_slide_structure, get_slide_structure_wcolor, delete_all_pptx_files, generate_pptx_from_text

# starting Fast API 
app = FastAPI() 
load_dotenv()
logger = logging.getLogger(__name__)
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")

# curl -X POST "http://localhost:5050/acra/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@CRA_SERVICE_CYBER.pptx"
@app.get("/acra/{folder_name}?add_info={add_info}")
async def summarize(folder_name: str, add_info: str = None):
    """
    Summarizes the content of PowerPoint files in a folder and updates a template PowerPoint file with the summary.
    The PowerPoint will be structured with a hierarchical format:
      - Main projects as headers
      - Subprojects under each main project
      - Information, alerts for each subproject
      - Events listed by service at the bottom

    Args:
        folder_name (str): The name of the folder containing PowerPoint files to analyze.
        additional_info (str, optional): Additional information or instructions for guiding the summarization process.

    Returns:
        dict: A dictionary containing the download URL of the updated PowerPoint file.

    Raises:
        HTTPException: If there's an error processing the PowerPoint files.
    """
    logger.info(f"Summarizing PPT for folder: {folder_name}")
    try:
        return summarize_ppt(folder_name, add_info)
    except Exception as e:
        # Log the exception for debugging
        print(f"Error in summarize_folder: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Summarize error: {str(e)}")

# curl -X POST "http://localhost:5050/acra/generate_report/CRA_SERVICE_CYBER" -H "accept: application/json" -H "Content-Type: application/json" -d "{\"info\": \"This is a test report\"}"
@app.post("/acra/generate_report/{folder_name}?info={info}")
async def generate_report(folder_name: str, info: str):
    """
    Génère un rapport à partir du texte fourni en utilisant une requête POST.
    
    Args:
        folder_name (str): Le nom du dossier où stocker le rapport
        info (str): Le texte à analyser pour générer le rapport
        
    Returns:
        dict: Résultat de la requête avec l'URL de téléchargement
    """
    try:
        return generate_pptx_from_text(folder_name, info)
    except Exception as e:
        print(f"Error in generate_report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Report generation error: {str(e)}")
    
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
@app.get("/download/{folder_name}/{filename}")
async def download_file(folder_name: str, filename: str):
    """
    Download a file from a specific folder on the server.
    This endpoint allows clients to download a file from a specific folder by specifying the folder name and filename in the URL path.
    
    Args:
        folder_name (str): The name of the folder containing the file.
        filename (str): The name of the file to be downloaded.
        
    Returns:
        FileResponse: A response containing the file to be downloaded.
        
    Raises:
        HTTPException: If the file does not exist, a 404 status code is returned with a detail message.
    """
    file_path = os.path.join(os.getenv("OUTPUT_FOLDER", "OUTPUT"), folder_name, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"File Not found at path: {file_path}")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

@app.delete("/delete_all_pptx_files/{foldername}")
async def delete_files(foldername:str):
    """
    Supprime tous les fichiers du dossier UPLOAD_FOLDER/foldername.

    Args:
        foldername (str): Le nom du dossier dans UPLOAD_FOLDER à vider.

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