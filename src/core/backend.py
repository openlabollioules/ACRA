import os
from dotenv import load_dotenv
import sys
from typing import Optional, Dict, Any # Added for type hinting
import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from services import update_table_with_project_data
from analist import analyze_presentation_with_colors, extract_projects_from_presentation
from .extract_and_summarize import aggregate_and_summarize, Generate_pptx_from_text

load_dotenv()
# Ensure UPLOAD_FOLDER and OUTPUT_FOLDER are absolute paths for consistency
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")) # Assuming backend.py is in src/core
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
if not os.path.isabs(UPLOAD_FOLDER):
    UPLOAD_FOLDER = os.path.join(BASE_DIR, UPLOAD_FOLDER)

OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")
if not os.path.isabs(OUTPUT_FOLDER):
    OUTPUT_FOLDER = os.path.join(BASE_DIR, OUTPUT_FOLDER)

def summarize_ppt(chat_id: str, add_info: Optional[str] = None, timestamp: Optional[str] = None, raw_structure_data: Optional[Dict[str, Any]] = None):
    """
    Summarizes content from PowerPoint files for a given chat_id or uses provided raw_structure_data.
    Then, updates a template PowerPoint file with this summarized data.

    Args:
        chat_id (str): The identifier for the conversation (previously folder_name).
        add_info (str, optional): Additional information to include in the summary prompt.
        timestamp (str, optional): Timestamp for unique filenames. Auto-generated if None.
        raw_structure_data (dict, optional): Pre-extracted project structure. If provided, file aggregation is skipped by aggregate_and_summarize.

    Returns:
        dict: Contains the filename and path to the summarized PowerPoint file, or an error structure.
    """
    
    # The core logic of file iteration is now inside aggregate_and_summarize if raw_structure_data is None.
    # Here, we directly call aggregate_and_summarize, passing all relevant parameters.
    print(f"Starting summarization for chat_id: {chat_id}")
    if raw_structure_data:
        print("summarize_ppt received raw_structure_data, will pass to aggregate_and_summarize.")

    # Call the updated aggregate_and_summarize function from extract_and_summarize.py
    # It handles using raw_structure_data or processing files from chat_id's folder.
    summarized_json_structure = aggregate_and_summarize(
        chat_id=chat_id, 
        add_info=add_info,
        timestamp=timestamp, # Pass timestamp along, though aggregate_and_summarize might not use it for logic
        raw_structure_data=raw_structure_data
    )

    # Validate the structure returned by aggregate_and_summarize
    if not isinstance(summarized_json_structure, dict) or "projects" not in summarized_json_structure:
        error_detail = summarized_json_structure.get("error", "Invalid structure from summarization") if isinstance(summarized_json_structure, dict) else "Unexpected response from summarization"
        log_message = f"ERROR: {error_detail} for chat_id {chat_id}."
        if isinstance(summarized_json_structure, dict) and summarized_json_structure.get("metadata", {}).get("errors"):
            log_message += f" Details: {summarized_json_structure['metadata']['errors']}"
        print(log_message)
        # Return an error structure compatible with CommandHandler expectations
        return {"error": log_message, "summary": None} 

    # Extract necessary data for PowerPoint generation
    project_data_for_pptx = summarized_json_structure.get("projects", {})
    upcoming_events_for_pptx = summarized_json_structure.get("upcoming_events", {})
    
    # Check if we have any actual data to put in the PowerPoint
    # The LLM might return an empty "projects" dict if it couldn't summarize anything meaningful.
    if not project_data_for_pptx and not upcoming_events_for_pptx: # If both are empty
        # Check if there were errors during the summarization process itself that didn't prevent a dict return
        metadata_errors = summarized_json_structure.get("metadata", {}).get("errors", [])
        source_file_errors = [sf.get("error") for sf in summarized_json_structure.get("source_files", []) if sf.get("error")]
        all_errors = metadata_errors + source_file_errors

        if all_errors:
            error_message = f"No project data to populate PowerPoint for chat {chat_id}. Errors encountered: {'; '.join(all_errors)}"
        else:
            # This case means aggregate_and_summarize ran, LLM ran, but LLM returned empty projects/events.
            # This might be valid if input files were empty or LLM deemed nothing summarizable.
            error_message = f"No summarizable project data or upcoming events found to populate PowerPoint for chat {chat_id}. The input might have been empty or non-relevant."
        
        print(f"WARNING: {error_message}")
        # We can still generate a blank or template-based PPTX, but it will be mostly empty.
        # Or, decide to return an error if an empty summary is not useful.
        # For now, let's proceed to generate a potentially empty PPTX but log the warning.
        # If an error should be returned, use: return {"error": error_message, "summary": None}

    # --- PowerPoint Generation --- # 
    # Create output directory for this chat_id if it doesn't exist
    # The output will be in OUTPUT_FOLDER/{chat_id}/summaries/
    chat_summary_output_dir = os.path.join(OUTPUT_FOLDER, chat_id, "summaries") # Specific subdirectory for summaries
    os.makedirs(chat_summary_output_dir, exist_ok=True)
    
    # Generate timestamp if not provided by caller (e.g. CommandHandler)
    current_timestamp = timestamp if timestamp else datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    output_filename = os.path.join(chat_summary_output_dir, f"summary_{chat_id}_{current_timestamp}.pptx")
    
    print(f"Creating summary PowerPoint at: {output_filename} for chat_id: {chat_id}")
    
    template_path = os.getenv("TEMPLATE_FILE", "templates/CRA_TEMPLATE_IA.pptx")
    if not os.path.isabs(template_path):
        template_path = os.path.join(BASE_DIR, template_path)

    if not os.path.exists(template_path):
        print(f"WARNING: Template file not found at {template_path}. update_table_with_project_data might fail or use a default.")
        # update_table_with_project_data should ideally handle template absence gracefully or take a Presentation object.

    # Call update_table_with_project_data to create/update the PowerPoint
    # This function is expected to handle template presence/absence and save to output_filename.
    try:
        final_ppt_path = update_table_with_project_data(
            pptx_path=template_path, 
            slide_index=0,      # Assuming first slide
            table_shape_index=0,      # Assuming first table on that slide
            project_data=project_data_for_pptx,
            output_path=output_filename,
            upcoming_events=upcoming_events_for_pptx
        )
        if "error" in final_ppt_path.lower(): # If update_table_with_project_data returns an error string
            print(f"Error from update_table_with_project_data for chat {chat_id}: {final_ppt_path}")
            return {"error": f"PPTX generation failed: {final_ppt_path}", "summary": None}

        created_filename = os.path.basename(final_ppt_path)
        print(f"Successfully created summary PowerPoint: {final_ppt_path} for chat_id: {chat_id}")
        return {"filename": created_filename, "summary": final_ppt_path} # Return path for upload

    except Exception as e:
        error_msg = f"Exception during PowerPoint generation for chat {chat_id}: {str(e)}"
        print(f"ERROR: {error_msg}")
        import traceback
        traceback.print_exc() # Print full traceback for debugging
        return {"error": error_msg, "summary": None}

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

