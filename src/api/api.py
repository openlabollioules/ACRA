import os,sys
import shutil
import uvicorn
import uuid
from pptx import presentation
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Query
from fastapi.responses import FileResponse  
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from analist import analyze_presentation , analyze_presentation_with_colors, extract_projects_from_presentation
from services import update_table_cell, update_table_multiple_cells, update_table_with_project_data
from core import aggregate_and_summarize, Generate_pptx_from_text
from OLLibrary.utils.log_service import setup_logging, get_logger
from dotenv import load_dotenv

# Set up logging
setup_logging(app_name="ACRA_API")
logger = get_logger(__name__)

# starting Fast API 
app = FastAPI() 
load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER")
logger.info(f"API starting with UPLOAD_FOLDER={UPLOAD_FOLDER}, OUTPUT_FOLDER={OUTPUT_FOLDER}")

# Configuration CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Permet toutes les origines en développement
    allow_credentials=True,
    allow_methods=["*"],  # Permet toutes les méthodes
    allow_headers=["*"],  # Permet tous les headers
)

# curl -X POST "http://localhost:5050/acra/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@CRA_SERVICE_CYBER.pptx"
@app.get("/acra/{folder_name}")
async def summarize_ppt(folder_name: str, additional_info: str = Query(None, description="Additional information or instructions for guiding the summarization process")):
    """
    Summarizes the content of PowerPoint files in a folder and updates a template PowerPoint file with the summary.
    The PowerPoint will be structured with 3 columns:
      - Column 1: Project names (each project in a separate row)
      - Column 2: Project information with colored text:
          * Black: Common information
          * Green: Advancements
          * Orange: Small alerts
          * Red: Critical alerts
      - Column 3: Upcoming events (all in one row)

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
        # Determine the target folder
        target_folder = UPLOAD_FOLDER
        if folder_name:
            target_folder = os.path.join(UPLOAD_FOLDER, folder_name)
        
        # Ensure the upload directory exists
        os.makedirs(target_folder, exist_ok=True)
        
        # Generate the summary from the PowerPoint files in the target folder
        logger.info(f"Generating summary for files in: {target_folder}")
        project_data = aggregate_and_summarize(target_folder, additional_info or "")
        
        # Check if we have any data to show
        if not project_data or (not project_data.get("activities") and not project_data.get("upcoming_events")):
            logger.warning(f"No data found in PowerPoint files for folder: {folder_name}")
            raise HTTPException(status_code=400, detail="Aucune information n'a pu être extraite des fichiers PowerPoint dans ce dossier.")
        
        # Generate a unique identifier for this summary
        unique_id = str(uuid.uuid4())[:8]  # Using first 8 characters of UUID for brevity
        
        # Set the output filename with the unique identifier
        if not folder_name:
            folder_name = "divers"
            
        os.makedirs(os.path.join(OUTPUT_FOLDER, folder_name), exist_ok=True)
        output_filename = f"{OUTPUT_FOLDER}/{folder_name}/summary_{unique_id}.pptx"
        
        # Update the template with the project data using the new format
        logger.info(f"Updating template with project data, output: {output_filename}")
        summarized_file_path = update_table_with_project_data(
            pptx_path=os.getenv("TEMPLATE_FILE"),  # Template file 
            slide_index=0,  # first slide
            table_shape_index=0,  # index of the table
            project_data=project_data,
            output_path=output_filename
        )

        # Return the download URL
        filename = os.path.basename(summarized_file_path)
        load_dotenv()
        download_url = f"http://localhost:5050/download/{OUTPUT_FOLDER}/{folder_name}/{filename}"
        logger.info(f"Summary generated successfully: {download_url}")
        return {"download_url": download_url}
    
    except Exception as e:
        # Log the exception for debugging
        logger.error(f"Error in summarize_folder: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Une erreur s'est produite: {str(e)}")

# Testing the function with : 
#  curl -X GET "http://localhost:5050/get_slide_structure/CRA_SERVICE_CYBER.pptx"
@app.get("/get_slide_structure/{foldername}")
async def get_slide_structure(foldername: str):
    """
    Analyse tous les fichiers PPTX présents dans le dossier pptx_folder.

    Returns:
        dict: Un dictionnaire contenant les structures des présentations analysées.

    Raises:
        HTTPException: Si aucun fichier PPTX n'est trouvé.
    """
    folder_path = os.path.join(UPLOAD_FOLDER,foldername)
    if not os.path.exists(folder_path):
        raise HTTPException(status_code=404, detail="Le dossier pptx_folder n'existe pas.")

    # Liste tous les fichiers dans le dossier
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith(".pptx")]

    # Si aucun fichier PPTX n'est trouvé, renvoyer un message
    if not pptx_files:
        return {"message": "Aucun fichier PPTX fourni."}

    # Analyse chaque fichier PPTX
    results = []
    for filename in pptx_files:
        file_path = os.path.join(folder_path, filename)
        
        try:            
            # Extraire les données sur les projets
            project_data = extract_projects_from_presentation(file_path)
            
            # Ajouter les deux ensembles de données au résultat
            results.append({
                "filename": filename, 
                "project_data": project_data
            })
        except Exception as e:
            results.append({"filename": filename, "error": f"Erreur lors de l'analyse: {str(e)}"})
    return {"presentations": results}


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
@app.get("/download/{base_folder}/{folder_name}/{filename}")
async def download_file(base_folder: str, folder_name: str, filename: str):
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
async def delete_all_pptx_files(foldername:str):
    """
    Supprime tous les fichiers du dossier UPLOAD_FOLDER/foldername.

    Args:
        foldername (str): Le nom du dossier dans UPLOAD_FOLDER à vider.

    Returns:
        dict: Un message confirmant la suppression des fichiers.
    
    Raises:
        HTTPException: Si le dossier n'existe pas.
    """
    pptx_folder_path = os.path.join(UPLOAD_FOLDER, foldername)
    
    if not os.path.exists(pptx_folder_path):
        return {"message": f"Le dossier {pptx_folder_path} n'existe pas."}

    # Liste des fichiers dans le dossier
    files = [f for f in os.listdir(pptx_folder_path) if os.path.isfile(os.path.join(pptx_folder_path, f))]
    
    if not files:
        return {"message": "Aucun fichier à supprimer."}

    # Suppression des fichiers un par un
    deleted_count = 0
    for file in files:
        file_path = os.path.join(pptx_folder_path, file)
        try:
            os.remove(file_path)
            deleted_count += 1
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{deleted_count} fichiers supprimés avec succès du dossier {pptx_folder_path}."}

@app.delete("/delete_output_files/{foldername}")
async def delete_output_files(foldername:str):
    """
    Supprime tous les fichiers du dossier OUTPUT/foldername.

    Args:
        foldername (str): Le nom du dossier dans OUTPUT à vider.

    Returns:
        dict: Un message confirmant la suppression des fichiers.
    
    Raises:
        HTTPException: Si le dossier n'existe pas.
    """
    output_folder_path = os.path.join(OUTPUT_FOLDER, foldername)
    
    if not os.path.exists(output_folder_path):
        return {"message": f"Le dossier {output_folder_path} n'existe pas."}

    # Liste des fichiers dans le dossier OUTPUT
    files = [f for f in os.listdir(output_folder_path) if os.path.isfile(os.path.join(output_folder_path, f))]
    
    if not files:
        return {"message": "Aucun fichier à supprimer dans le dossier OUTPUT."}

    # Suppression des fichiers un par un
    deleted_count = 0
    for file in files:
        file_path = os.path.join(output_folder_path, file)
        try:
            os.remove(file_path)
            deleted_count += 1
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{deleted_count} fichiers supprimés avec succès du dossier {pptx_folder_path}."}

@app.delete("/delete_output_files/{foldername}")
async def delete_output_files(foldername:str):
    """
    Supprime tous les fichiers du dossier OUTPUT/foldername.

    Args:
        foldername (str): Le nom du dossier dans OUTPUT à vider.

    Returns:
        dict: Un message confirmant la suppression des fichiers.
    
    Raises:
        HTTPException: Si le dossier n'existe pas.
    """
    output_folder_path = os.path.join(OUTPUT_FOLDER, foldername)
    
    if not os.path.exists(output_folder_path):
        return {"message": f"Le dossier {output_folder_path} n'existe pas."}

    # Liste des fichiers dans le dossier OUTPUT
    files = [f for f in os.listdir(output_folder_path) if os.path.isfile(os.path.join(output_folder_path, f))]
    
    if not files:
        return {"message": "Aucun fichier à supprimer dans le dossier OUTPUT."}

    # Suppression des fichiers un par un
    deleted_count = 0
    for file in files:
        file_path = os.path.join(output_folder_path, file)
        try:
            os.remove(file_path)
            deleted_count += 1
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{deleted_count} fichiers supprimés avec succès du dossier OUTPUT/{foldername}."}
@app.post("/generate_text_report/{foldername}")
async def generate_text_report(foldername: str, text_data: dict):
    """Takes the ACRA Info and generates a PPTX from text files, following the template"""
    try:
        # Extract text info from request body
        info = text_data.get("info", "")
        
        # Determine the target folder
        target_folder = UPLOAD_FOLDER
        if foldername:
            target_folder = os.path.join(UPLOAD_FOLDER, foldername)
        
        # Ensure the upload directory exists
        os.makedirs(target_folder, exist_ok=True)
        
        # Generate the summary using our updated function with the text information
        project_data = Generate_pptx_from_text(target_folder, info)
        
        # Check if we have any data to show
        if not project_data or (not project_data.get("activities") and not project_data.get("upcoming_events")):
            raise HTTPException(status_code=400, detail="Aucune information n'a pu être extraite du texte fourni.")
        
        # Set the output filename
        output_filename = f"{OUTPUT_FOLDER}/updated_presentation_from_text.pptx"
        if foldername:
            output_filename = f"{OUTPUT_FOLDER}/{foldername}_text_summary.pptx"
        
        # Update the template with the project data using the new format
        generated_pptx = update_table_with_project_data(
            pptx_path=os.getenv("TEMPLATE_FILE"),  # Template file 
            slide_index=0,  # first slide
            table_shape_index=0,  # index of the table
            project_data=project_data,
            output_path=output_filename
        )

        # Return the download URL
        filename = os.path.basename(generated_pptx)
        # TODO: add a way to download a file from the UPLOAD_FOLDER
        # copy the file to the upload folder
        shutil.copy(output_filename, os.path.join(target_folder, os.path.basename(output_filename)))

        return {"download_url": f"http://localhost:5050/download/pptx_folder/{foldername}/{filename}"}
    
    except Exception as e:
        # Log the exception for debugging
        print(f"Error in generate_text_report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Une erreur s'est produite: {str(e)}")

def run():
    uvicorn.run(app, host="0.0.0.0", port=5050)


if __name__ == "__main__":
    summarize_ppt("669fa53b-649a-4023-8066-4cd86670e88b")