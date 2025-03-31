import os,sys
import shutil
import uvicorn
from pptx import presentation
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Query
from fastapi.responses import FileResponse  
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from analist import analyze_presentation , analyze_presentation_with_colors, extract_projects_from_presentation
from services import update_table_cell, update_table_multiple_cells, update_table_with_project_data
from core import aggregate_and_summarize

# starting Fast API 
app = FastAPI() 
load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER")

# curl -X POST "http://localhost:5050/acra/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@CRA_SERVICE_CYBER.pptx"
@app.get("/acra/{folder_name}")
async def summarize_ppt(folder_name: str):
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

    Returns:
        dict: A dictionary containing the download URL of the updated PowerPoint file.

    Raises:
        HTTPException: If there's an error processing the PowerPoint files.
    """
    try:
        # Determine the target folder
        target_folder = UPLOAD_FOLDER
        if folder_name:
            target_folder = os.path.join(UPLOAD_FOLDER, folder_name)
        
        # Ensure the upload directory exists
        os.makedirs(target_folder, exist_ok=True)
        
        # Generate the summary from the PowerPoint files in the target folder
        project_data = aggregate_and_summarize(target_folder)
        
        # Check if we have any data to show
        if not project_data or (not project_data.get("activities") and not project_data.get("upcoming_events")):
            raise HTTPException(status_code=400, detail="Aucune information n'a pu être extraite des fichiers PowerPoint dans ce dossier.")
        
        # Set the output filename
        output_filename = f"{OUTPUT_FOLDER}/updated_presentation.pptx"
        if folder_name:
            output_filename = f"{OUTPUT_FOLDER}/{folder_name}_summary.pptx"
        
        # Update the template with the project data using the new format
        summarized_file_path = update_table_with_project_data(
            pptx_path=os.getenv("TEMPLATE_FILE"),  # Template file 
            slide_index=0,  # first slide
            table_shape_index=0,  # index of the table
            project_data=project_data,
            output_path=output_filename
        )

        # Return the download URL
        filename = os.path.basename(summarized_file_path)
        return {"download_url": f"http://localhost:5050/download/{filename}"}
    
    except Exception as e:
        # Log the exception for debugging
        print(f"Error in summarize_folder: {str(e)}")
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
async def delete_all_pptx_files(foldername:str):
    """
    Supprime tous les fichiers du dossier pptx_folder.

    Returns:
        dict: Un message confirmant la suppression des fichiers.
    
    Raises:
        HTTPException: Si le dossier n'existe pas.
    """
    pptx_folder = os.path.join(UPLOAD_FOLDER, foldername)
    if not os.path.exists(pptx_folder):
        raise HTTPException(status_code=404, detail="Le dossier pptx_folder n'existe pas.")

    # Liste des fichiers dans le dossier
    files = os.listdir(pptx_folder)
    
    if not files:
        return {"message": "Aucun fichier à supprimer."}

    # Suppression des fichiers un par un
    for file in files:
        file_path = os.path.join(pptx_folder, file)
        try:
            os.remove(file_path)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{len(files)} fichiers supprimés avec succès."}


def run():
    uvicorn.run(app, host="0.0.0.0", port=5050)


if __name__ == "__main__":
    summarize_ppt("669fa53b-649a-4023-8066-4cd86670e88b")