import os
import shutil
from dotenv import load_dotenv
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from services import update_table_with_project_data
from analist import  analyze_presentation_with_colors, extract_projects_from_presentation
from .extract_and_summarize import aggregate_and_summarize, Generate_pptx_from_text

load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")

def summarize_ppt(folder_name : str, add_info : str = None, timestamp : str = None):
        """
    Summarizes the content of PowerPoint files in a folder and updates a template PowerPoint file with the summary.
    The PowerPoint will be structured with a hierarchical format:
      - Main projects as headers
      - Subprojects under each main project
      - Information, alerts for each subproject
      - Events listed by service at the bottom

    Args:
        folder_name (str): The name of the folder containing PowerPoint files to analyze.
        add_info (str, optional): Additional information to include in the summary.
        timestamp (str, optional): Timestamp to use in the filename for uniqueness. If None, will use the current time.

    Returns:
        dict: A dictionary containing the download URL of the updated PowerPoint file.
    """
    # Determine the target folder
        target_folder = UPLOAD_FOLDER
        if folder_name:
            target_folder = os.path.join(UPLOAD_FOLDER, folder_name)
        
        # Ensure the upload directory exists
        os.makedirs(target_folder, exist_ok=True)
        
        print(f"Starting summarization for folder: {target_folder}")
        
        # List files in folder for diagnostics
        files_in_folder = os.listdir(target_folder)
        pptx_files = [f for f in files_in_folder if f.lower().endswith(".pptx")]
        print(f"Found {len(pptx_files)} PPTX files in folder: {pptx_files}")
        
        if not pptx_files:
            raise Exception(f"Aucun fichier PowerPoint (.pptx) trouvé dans le dossier {folder_name}.")
        
        # Récupérer les données des projets à partir de get_slide_structure
        structure_result = aggregate_and_summarize(folder_name, add_info)
        project_data = structure_result.get("projects", {})
        upcoming_events = structure_result.get("upcoming_events", {})
        errors = structure_result.get("metadata", {}).get("errors", [])
        
        # Print diagnostic information
        print(f"Project data contains {len(project_data)} top-level projects")
        print(f"Upcoming events contains data for {len(upcoming_events)} services")
        
        # Check if we have any data to show
        if not project_data or len(project_data) == 0:
            error_message = "Aucune information n'a pu être extraite des fichiers PowerPoint dans ce dossier."
            if errors:
                error_message += f" Erreurs rencontrées: {'; '.join(errors)}"
            print(f"ERROR: {error_message}")
            raise Exception(error_message)
        
        # Create output directory for this folder
        folder_output_path = os.path.join(OUTPUT_FOLDER, folder_name)
        os.makedirs(folder_output_path, exist_ok=True)
        
        # Generate timestamp if not provided
        if timestamp is None:
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Set the output filename directly in the subfolder, including timestamp
        output_filename = os.path.join(folder_output_path, f"{folder_name}_summary_{timestamp}.pptx")
        
        print(f"Creating summary PowerPoint at: {output_filename}")
        
        # Update the template with the project data using the new format
        summarized_file_path = update_table_with_project_data(
            pptx_path=os.getenv("TEMPLATE_FILE", "templates/CRA_TEMPLATE_IA.pptx"),  # Template file 
            slide_index=0,  # first slide
            table_shape_index=0,  # index of the table
            project_data=project_data,
            output_path=output_filename,
            upcoming_events=upcoming_events  # Pass upcoming events by service
        )

        # Return the download URL
        filename = os.path.basename(summarized_file_path)
        return {"filename": filename, "summary": summarized_file_path}

def get_slide_structure(foldername : str):
    # Check if foldername is None
    if foldername is None:
        raise Exception("Le nom du dossier (foldername) ne peut pas être None.")
        
    folder_path = os.path.join(UPLOAD_FOLDER, foldername)
    if not os.path.exists(folder_path):
        raise Exception("Le dossier n'existe pas.")

    # Liste tous les fichiers dans le dossier
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith(".pptx")]

    # Si aucun fichier PPTX n'est trouvé, renvoyer un message
    if not pptx_files:
        return {"message": "Aucun fichier PPTX fourni."}
    
    # Fonction récursive pour fusionner des dictionnaires de projets hiérarchiques
    def merge_project_dictionaries(dict1, dict2):
        result = dict1.copy()
        
        for key, value in dict2.items():
            if key in result:
                # Si la clé existe dans les deux dictionnaires
                if isinstance(value, dict) and isinstance(result[key], dict):
                    # Si les deux valeurs sont des dictionnaires, fusion récursive
                    if "information" in value and "information" in result[key]:
                        # C'est un niveau terminal, fusionner les champs
                        result[key]["information"] += "\n\n" + value["information"] if result[key]["information"] else value["information"]
                        result[key]["critical"].extend([item for item in value.get("critical", []) if item not in result[key]["critical"]])
                        result[key]["small"].extend([item for item in value.get("small", []) if item not in result[key]["small"]])
                        result[key]["advancements"].extend([item for item in value.get("advancements", []) if item not in result[key]["advancements"]])
                    else:
                        # C'est un niveau intermédiaire, fusionner récursivement
                        result[key] = merge_project_dictionaries(result[key], value)
                else:
                    # Si les types sont différents, garder celui de dict2
                    result[key] = value
            else:
                # Si la clé n'existe pas dans dict1, l'ajouter
                result[key] = value
                
        return result
    
    # Fonction pour extraire le nom du service à partir du nom de fichier
    def extract_service_name(filename):
        # Extraire le titre après le dernier underscore
        parts = filename.split('_')
        if len(parts) > 1:
            # Prendre les parties après le premier élément qui contient généralement l'ID
            title_parts = parts[1:]
            # Recombiner en supprimant l'extension .pptx
            title = ' '.join(title_parts).replace('.pptx', '').strip()
            # Capitaliser correctement
            title = ' '.join(word.capitalize() for word in title.split())
            return title
        else:
            # Si pas d'underscore, on retire juste l'extension
            return filename.replace('.pptx', '').strip()

    # Analyse chaque fichier PPTX et fusionne les données
    all_projects = {}
    upcoming_events_by_service = {}
    processed_files = []
    
    for filename in pptx_files:
        file_path = os.path.join(folder_path, filename)
        
        try:            
            # Extraire les données sur les projets avec le nouveau format
            project_data = extract_projects_from_presentation(file_path)
            
            # Extraire le nom du service
            service_name = extract_service_name(filename)
            
            # Récupérer les événements collectés depuis les métadonnées
            collected_events = []
            if "metadata" in project_data and "collected_upcoming_events" in project_data["metadata"]:
                collected_events = project_data["metadata"]["collected_upcoming_events"]
            
            # Ajouter les événements au service correspondant
            if collected_events:
                if service_name not in upcoming_events_by_service:
                    upcoming_events_by_service[service_name] = []
                for event in collected_events:
                    if event not in upcoming_events_by_service[service_name]:
                        upcoming_events_by_service[service_name].append(event)
            
            # Fusionner les projets
            if "projects" in project_data:
                all_projects = merge_project_dictionaries(all_projects, project_data["projects"])
            
            # Ajouter le fichier à la liste des présentations traitées
            processed_files.append({
                "filename": filename,
                "service_name": service_name,
                "processed": True,
                "events_count": len(collected_events)
            })
        except Exception as e:
            processed_files.append({"filename": filename, "error": f"Erreur lors de l'analyse: {str(e)}"})
    
    # Si aucun événement n'a été trouvé, ajouter un message par défaut pour chaque service traité
    if not upcoming_events_by_service:
        for file_info in processed_files:
            if "service_name" in file_info and "processed" in file_info and file_info["processed"]:
                service_name = file_info["service_name"]
                if service_name not in upcoming_events_by_service:
                    upcoming_events_by_service[service_name] = ["Aucun événement particulier prévu pour ce service."]
    
    # Créer la structure finale
    return {
        "projects": all_projects,
        "upcoming_events": upcoming_events_by_service,
        "metadata": {
            "processed_files": len(processed_files),
            "folder": foldername
        },
        "source_files": processed_files
    }

def get_slide_structure_wcolor(filename : str):
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    if not os.path.exists(file_path):
        raise Exception("File not found")

    slides_data = analyze_presentation_with_colors(file_path)
    return {"filename": filename, "slide data": slides_data}

def delete_all_pptx_files(foldername : str):
    pptx_folder = os.path.join(UPLOAD_FOLDER, foldername)
    if not os.path.exists(pptx_folder):
        raise Exception("Le dossier pptx_folder n'existe pas.")

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
            raise Exception(f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{len(files)} fichiers supprimés avec succès."}

def generate_pptx_from_text(foldername : str, info : str, timestamp : str = None):
    """
    Takes the ACRA Info and generates a PPTX from text files, following the template
    
    Args:
        foldername (str): Name of the folder to store the PowerPoint
        info (str): Text to analyze for the PowerPoint generation
        timestamp (str, optional): Timestamp to use in the filename for uniqueness. If None, will use the current time.
    
    Returns:
        dict: A dictionary containing filename and path to the generated PowerPoint
    """
    
    # Determine the target folder
    target_folder = UPLOAD_FOLDER
    if foldername:
        target_folder = os.path.join(UPLOAD_FOLDER, foldername)
    
    # Ensure the upload directory exists
    os.makedirs(target_folder, exist_ok=True)
    
    # Generate the summary using our updated function with the text information
    project_data = Generate_pptx_from_text(target_folder, info)
    print("project_data :", project_data)

    # Ensure project_data is a dictionary, not a list
    if isinstance(project_data, dict):
        # Check if we need to extract the 'projects' key
        if 'projects' in project_data:
            projects = project_data.get('projects', {})
            upcoming = project_data.get('upcoming_events', {})
        else:
            projects = project_data
            upcoming = {}
    else:
        # If it's not a dictionary, create an empty one
        print("Warning: project_data is not a dictionary. Creating empty structure.")
        projects = {}
        upcoming = {}
    
    # Generate timestamp if not provided
    if timestamp is None:
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Save the file in pptx_folder/chat_id instead of OUTPUT/chat_id
    generated_filename = f"{foldername}_text_summary_{timestamp}.pptx"
    output_filename = os.path.join(target_folder, generated_filename)
    print(f"Creating text-generated PowerPoint at: {output_filename}")
    
    # Update the template with the project data using the new format
    load_dotenv()
    generated_pptx = update_table_with_project_data(
        pptx_path=os.getenv("TEMPLATE_FILE", "templates/CRA_TEMPLATE_IA.pptx"),  # Template file 
        slide_index=0,  # first slide
        table_shape_index=0,  # index of the table
        project_data=projects,
        output_path=output_filename,
        upcoming_events=upcoming
    )

    filename = os.path.basename(generated_pptx)
    return {"filename": filename, "summary": generated_pptx}

