import os,sys
import re
import json
from langchain_core.prompts import PromptTemplate
from dotenv import load_dotenv
import time
from typing import Optional, Dict, Any, List

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Load environment variables
load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
# Ensure UPLOAD_FOLDER is absolute for consistent path resolution
if not os.path.isabs(UPLOAD_FOLDER):
    BASE_DIR_FOR_UPLOAD = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")) # Assuming this file is in src/core
    UPLOAD_FOLDER = os.path.join(BASE_DIR_FOR_UPLOAD, UPLOAD_FOLDER)

from langchain_ollama import OllamaLLM
summarize_model = OllamaLLM(model="qwen3:30b-a3b", base_url="http://host.docker.internal:11434", temperature=0.7, num_ctx=132000)

from analist import extract_projects_from_presentation
from OLLibrary.utils.text_service import remove_tags_no_keep

def extract_common_and_upcoming_info(project_data):
    """
    Extract common information, upcoming work information, and alerts from project data.
    
    Parameters:
      project_data (dict): Project data dictionary extracted from presentations.
    
    Returns:
      dict: A dictionary containing common_info, upcoming_info, advancements, small_alerts, and critical_alerts
    """
    common_info = []
    upcoming_info = ""
    advancements = []
    small_alerts = []
    critical_alerts = []
    
    # Extract common information from all projects
    for project_name, project_info in project_data.items():
        if project_name == "metadata" or project_name == "upcoming_events":
            continue
            
        if "information" in project_info:
            info_text = project_info["information"]
            
            # Check if the information contains details about upcoming week
            upcoming_week_match = re.search(r"Evénements de la semaine à venir(.*?)$", info_text, re.DOTALL)
            if upcoming_week_match:
                # Split the information: before match goes to common_info, the match itself goes to upcoming_info
                common_part = info_text[:upcoming_week_match.start()]
                upcoming_part = upcoming_week_match.group(1).strip()
                
                if common_part:
                    common_info.append(f"{project_name}: {common_part}")
                if upcoming_part:
                    upcoming_info += f"{project_name}: {upcoming_part}\n"
            else:
                common_info.append(f"{project_name}: {info_text}")
        
        # Process alerts and store them in separate categories
        if "alerts" in project_info:
            alerts = project_info["alerts"]
            
            if alerts.get("advancements"):
                for advancement in alerts.get("advancements", []):
                    advancements.append(f"{project_name}: {advancement}")
            
            if alerts.get("small_alerts"):
                for alert in alerts.get("small_alerts", []):
                    small_alerts.append(f"{project_name}: {alert}")
            
            if alerts.get("critical_alerts"):
                for alert in alerts.get("critical_alerts", []):
                    critical_alerts.append(f"{project_name}: {alert}")
    
    # Add upcoming events from project_data if available
    if "upcoming_events" in project_data:
        upcoming_info += project_data["upcoming_events"]
    
    # Prepare the result dictionary
    result = {
        "common_info": "\n\n".join(common_info),
        "upcoming_info": upcoming_info if upcoming_info else "Aucun événement particulier prévu pour la semaine à venir.",
        "advancements": "\n".join(advancements) if advancements else "Aucun avancement significatif à signaler.",
        "small_alerts": "\n".join(small_alerts) if small_alerts else "Aucune alerte mineure à signaler.",
        "critical_alerts": "\n".join(critical_alerts) if critical_alerts else "Aucune alerte critique à signaler."
    }
    
    return result

def aggregate_and_summarize(chat_id: str, add_info: Optional[str] = None, timestamp: Optional[str] = None, raw_structure_data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Aggregates information from PPTX files or uses provided raw structure, then summarizes using an LLM.

    Parameters:
      chat_id (str): Identifier for the chat/conversation, used to locate files if raw_structure_data is not provided.
      add_info (str, optional): Additional information to include in the LLM summarization prompt.
      timestamp (str, optional): Timestamp, currently not used by this function but kept for signature consistency.
      raw_structure_data (dict, optional): Pre-extracted project structure. If provided, file processing is skipped.
    
    Returns:
      dict: LLM-summarized project data in JSON format, or the raw aggregated data if summarization fails.
    """
    final_data_for_llm: Dict[str, Any]
    extraction_errors: List[str] = []
    processed_files_metadata: List[Dict[str, Any]] = []
    file_count = 0

    # Path 1: Use provided raw_structure_data if valid
    if raw_structure_data and isinstance(raw_structure_data, dict) and \
       ("projects" in raw_structure_data or "upcoming_events" in raw_structure_data): # Check for essential data keys
        print(f"Using provided raw_structure_data for chat_id: {chat_id}")
        final_data_for_llm = json.loads(json.dumps(raw_structure_data)) # Deep copy to avoid modifying the cache

        # Ensure essential keys exist and populate metadata from the raw_structure_data
        if "projects" not in final_data_for_llm: final_data_for_llm["projects"] = {}
        if "upcoming_events" not in final_data_for_llm: final_data_for_llm["upcoming_events"] = {}
        
        # Metadata handling from cache
        cached_metadata = final_data_for_llm.get("metadata", {})
        file_count = cached_metadata.get("processed_files", 0)
        # If processed_files count is zero or missing, try to infer from source_files
        if not file_count and "source_files" in final_data_for_llm and isinstance(final_data_for_llm["source_files"], list):
            file_count = len([sf for sf in final_data_for_llm["source_files"] if isinstance(sf, dict) and sf.get("processed")])
        
        processed_files_metadata = final_data_for_llm.get("source_files", [])
        # Ensure processed_files_metadata is a list of dicts
        if not isinstance(processed_files_metadata, list) or not all(isinstance(item, dict) for item in processed_files_metadata):
            print(f"Warning: 'source_files' in raw_structure_data for chat {chat_id} is not a list of dicts. Resetting.")
            processed_files_metadata = [] # Reset if format is incorrect

        extraction_errors = cached_metadata.get("errors", [])
        if not isinstance(extraction_errors, list): extraction_errors = []


        # Ensure the final metadata structure is consistent
        final_data_for_llm["metadata"] = {
            "folder": cached_metadata.get("folder", chat_id), # Use cached folder or current chat_id
            "processed_files": file_count,
            "errors": extraction_errors,
        }
        final_data_for_llm["source_files"] = processed_files_metadata

    # Path 2: No valid raw_structure_data, process files from folder
    else:
        if raw_structure_data: # It was provided but invalid
             print(f"Provided raw_structure_data for chat_id {chat_id} was invalid or empty. Processing files from folder.")
        else: # It was not provided at all
             print(f"No raw_structure_data provided for chat_id: {chat_id}. Processing files from folder.")

        full_path = os.path.join(UPLOAD_FOLDER, chat_id)
        print(f"Processing folder for file aggregation: {full_path}")

        if not os.path.exists(full_path) or not os.path.isdir(full_path):
            error_msg = f"Error: Folder {full_path} does not exist or is not a directory for chat_id {chat_id}."
            print(error_msg)
            extraction_errors.append(error_msg)
            return {
                "projects": {}, "upcoming_events": {},
                "metadata": {"processed_files": 0, "folder": chat_id, "errors": extraction_errors},
                "source_files": []
            }
        
        all_files_in_dir = os.listdir(full_path)
        pptx_files = [f for f in all_files_in_dir if f.lower().endswith(".pptx")]
        print(f"PPTX files found in {full_path}: {pptx_files}")
        
        if not pptx_files:
            error_msg = f"No PowerPoint files found in folder {full_path} for chat_id {chat_id}."
            print(error_msg)
            # extraction_errors.append(error_msg) # No error if folder is just empty
            return {
                "projects": {}, "upcoming_events": {},
                "metadata": {"processed_files": 0, "folder": chat_id, "errors": extraction_errors}, # errors list might be empty
                "source_files": []
            }
        
        current_aggregated_projects: Dict[str, Any] = {}
        current_aggregated_events: Dict[str, List[str]] = {}

        for filename in pptx_files:
            file_path = os.path.join(full_path, filename)
            print(f"Processing file for aggregation: {file_path}")
            
            try:
                # extract_projects_from_presentation is from analist module
                file_project_data = extract_projects_from_presentation(file_path) 
                file_count += 1
                
                service_name_parts = os.path.basename(filename).replace('.pptx', '').split('_')
                service_name = service_name_parts[-1] if len(service_name_parts) > 1 else service_name_parts[0]
                
                processed_file_info = {"filename": filename, "service_name": service_name, "processed": True}
                
                if "projects" in file_project_data and file_project_data["projects"]:
                    project_count_in_file = len(file_project_data["projects"])
                    processed_file_info["project_count"] = project_count_in_file
                    # Basic merge: For a more robust solution, a deep merge function would be better
                    for main_project_name, main_project_content in file_project_data["projects"].items():
                        if main_project_name not in current_aggregated_projects:
                            current_aggregated_projects[main_project_name] = main_project_content
                        else: # Rudimentary merge for top-level project data
                            if isinstance(main_project_content, dict) and isinstance(current_aggregated_projects[main_project_name], dict):
                                for key in ["information", "critical", "small", "advancements"]:
                                    if key in main_project_content:
                                        if key == "information" and current_aggregated_projects[main_project_name].get(key) and main_project_content[key]:
                                            current_aggregated_projects[main_project_name][key] += "\n" + main_project_content[key]
                                        elif key == "information" and main_project_content[key]:
                                             current_aggregated_projects[main_project_name][key] = main_project_content[key]
                                        elif key != "information": # For lists like critical, small, advancements
                                            current_aggregated_projects[main_project_name].setdefault(key, []).extend(
                                                item for item in main_project_content[key] if item not in current_aggregated_projects[main_project_name].get(key,[])
                                            )
                else: # No projects found in this file
                    error_detail = file_project_data.get("metadata", {}).get("error", f"No projects extracted from {filename}")
                    print(f"Warning/Error in file {filename}: {error_detail}")
                    # Only add to extraction_errors if it's a genuine error from the extractor
                    if "error" in file_project_data.get("metadata", {}):
                        extraction_errors.append(f"File {filename}: {error_detail}")
                    processed_file_info["warning" if "error" not in file_project_data.get("metadata", {}) else "error"] = error_detail
                
                processed_files_metadata.append(processed_file_info)
                
                # Handle upcoming events from file's metadata
                if "metadata" in file_project_data and "collected_upcoming_events" in file_project_data["metadata"]:
                    events = file_project_data["metadata"]["collected_upcoming_events"]
                    if events and isinstance(events, list):
                        processed_file_info["events_count"] = len(events)
                        current_aggregated_events.setdefault(service_name, []).extend(
                            event for event in events if event not in current_aggregated_events.get(service_name, [])
                        )
                    else: 
                        processed_file_info["events_count"] = 0
            except Exception as e:
                error_message = f"Exception processing file {filename}: {str(e)}"
                print(error_message, exc_info=True)
                extraction_errors.append(error_message)
                processed_files_metadata.append({"filename": filename, "service_name": "Unknown", "processed": False, "error": str(e)})
        
        if file_count > 0 and not current_aggregated_projects and not current_aggregated_events:
            msg = f"Warning: {file_count} files processed for chat {chat_id}, but no project data or events were aggregated."
            print(msg)
            # Not necessarily an error if files were empty/irrelevant
        
        if file_count == 0 and pptx_files: # Files existed but none could be processed
            error_msg = f"No files were successfully processed in {full_path} for chat {chat_id}, though PPTX files were present."
            print(error_msg)
            extraction_errors.append(error_msg)
            # Return early as there's nothing to summarize
            return {"projects": {}, "upcoming_events": {}, "metadata": {"processed_files": 0, "folder": chat_id, "errors": extraction_errors}, "source_files": processed_files_metadata}

        final_data_for_llm = {
            "projects": current_aggregated_projects,
            "upcoming_events": current_aggregated_events,
            "metadata": {"processed_files": file_count, "folder": chat_id, "errors": extraction_errors},
            "source_files": processed_files_metadata
        }

    # LLM Summarization Part (common for both paths if data exists)
    
    # Log structure before LLM
    print(f"Data for LLM (chat {chat_id}): {len(final_data_for_llm.get('projects', {}))} projects, {len(final_data_for_llm.get('upcoming_events', {}))} services with events.")
    if not final_data_for_llm.get("projects") and not final_data_for_llm.get("upcoming_events"):
        print(f"Warning: No project data or upcoming events to send to LLM for chat {chat_id}. This might be intended if input was empty.")

    prompt_inputs = {
        "project_data": json.dumps(final_data_for_llm, indent=2, ensure_ascii=False),
        "temp_add_info": ""
    }
    if add_info:
        prompt_inputs["temp_add_info"] = f"Voici des informations supplémentaires qui peuvent être utiles pour la synthèse: {add_info}"
    
    summarization_template = """    Tu es un assistant chargé de résumer des informations de projets et de les formater.

    Voici les données des projets:
    {project_data}
    
    Analyse ces données et identifie les points clés pour chaque projet et sous-projet.
    Pour chaque entrée, tu peux conserver la structure mais synthétise les informations
    pour qu'elles soient plus concises tout en préservant les détails importants.
    
    IMPORTANT: Inclus TOUTES les informations dans le champ "information" de chaque projet. MAIS quand tu identifies une information comme étant un avancement, une alerte mineure ou une alerte critique, COPIE-LA ÉGALEMENT dans la catégorie correspondante (critical, small, advancements) pour qu'elle puisse être colorée. Ainsi, le texte apparaîtra dans le champ information mais sera automatiquement coloré.
    
    CONCERNANT LES ÉVÉNEMENTS À VENIR: Ne conserve dans la section "upcoming_events" QUE les informations qui sont EXPLICITEMENT des événements futurs. Si un élément ne mentionne pas clairement un événement à venir, retire-le de cette section. Si aucun événement futur n'est clairement identifié, laisse la section "upcoming_events" VIDE avec un objet vide {{}}.
    
    Les alertes critiques, alertes mineures et avancements doivent être conservés tels quels,
    mais tu peux éliminer les redondances éventuelles. Soit vraiment le plus concis possible mais il faut également
    pouvoir retransmettre le maximum d'informations. N'hésites pas à synthétiser en quelques mots (essaie de te contenir à 10 mots environs)
    mais il ne faut pas perdre d'informations importantes.
    
    {temp_add_info}

    Réponds uniquement avec la structure JSON modifiée, sans texte d'introduction ni d'explication.
    """
    
    prompt = summarization_template.format(**prompt_inputs)
    
    # Skip LLM if there's truly nothing to summarize (empty projects AND empty events)
    if not final_data_for_llm.get("projects") and not final_data_for_llm.get("upcoming_events"):
        print(f"Skipping LLM summarization for chat {chat_id} as no project or event data was found/aggregated.")
        # Return the (likely empty) structure with its metadata
        return final_data_for_llm 
    
    try:
        prompt_size = len(prompt.encode('utf-8'))
        print(f"Summarization prompt size: {prompt_size} bytes for chat {chat_id}")
        # Adjusted warning thresholds
        if prompt_size > 150000: 
            print(f"WARNING: Very large prompt detected ({prompt_size} bytes) for chat {chat_id}, LLM may timeout or fail.")
        elif prompt_size < 200 and not (final_data_for_llm.get("projects") or final_data_for_llm.get("upcoming_events")): # Very small prompt AND no actual data
             print(f"Warning: Very small prompt size ({prompt_size} bytes) for chat {chat_id} and no project/event data. Likely empty input. Skipping LLM.")
             return final_data_for_llm

        print(f"Calling LLM for summarization for chat {chat_id}...")
        llm_response = summarize_model.invoke(prompt)
        print(f"LLM response received successfully for chat {chat_id}")
        
        llm_response_cleaned = remove_tags_no_keep(llm_response, "<think>", "</think>")
        # Attempt to find JSON block, otherwise assume the whole response is JSON
        json_match = re.search(r'```json\s*(.*?)```', llm_response_cleaned, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = llm_response_cleaned # Assume entire cleaned response is the JSON
        
        json_str = json_str.strip()
        summarized_result = json.loads(json_str)
        
        print(f"LLM summarization completed successfully for chat {chat_id}")

        # Ensure original/current metadata and source_files are preserved or merged into LLM response
        # The LLM should ideally return the full structure including "metadata" and "source_files".
        # If it doesn't, we merge them back to ensure consistency.
        if "metadata" not in summarized_result:
            summarized_result["metadata"] = final_data_for_llm.get("metadata", {})
        else: # Merge, giving precedence to original metadata if keys conflict, or update if LLM adds new info
            original_meta = final_data_for_llm.get("metadata", {})
            for k, v in original_meta.items():
                summarized_result["metadata"].setdefault(k, v)
        
        if "source_files" not in summarized_result:
            summarized_result["source_files"] = final_data_for_llm.get("source_files", [])
            
        return summarized_result
        
    except json.JSONDecodeError as json_e:
        print(f"JSON Decode Error during LLM summarization for chat {chat_id}: {str(json_e)}")
        print(f"LLM Response (cleaned) that caused error: '{json_str[:500]}...'")
        final_data_for_llm.setdefault("metadata", {}).setdefault("errors", []).append(f"LLM JSON Decode Error: {str(json_e)}")
        return final_data_for_llm # Return raw aggregated data
    except Exception as e:
        print(f"Error during LLM summarization for chat {chat_id}: {str(e)}", exc_info=True)
        if "EOF" in str(e) or "Connection" in str(e) or "timeout" in str(e).lower():
            print("Connection error detected - likely Ollama service issue or timeout.")
        
        print(f"Returning raw aggregated data as fallback for chat {chat_id} due to LLM error.")
        final_data_for_llm.setdefault("metadata", {}).setdefault("errors", []).append(f"LLM summarization failed: {str(e)}")
        return final_data_for_llm

def Generate_pptx_from_text(chat_id: str, info: Optional[str] = None, timestamp: Optional[str] = None) -> Dict[str, Any]: 
    """
    Generate a JSON structure from text input that can be used by update_table_with_project_data.
    Uses an LLM to process the text information and return it in the specified JSON format.
    
    Parameters:
      chat_id (str): Identifier for the chat/conversation.
      info (str, optional): Text information to process and structure into JSON format.
      timestamp (str, optional): Timestamp, currently not used by this function but kept for signature consistency.

    Returns:
      dict: A dictionary with project data in the specified JSON format.
    """
    if not info:
        return {
            "projects": {},
            "upcoming_events": {},
            "metadata": {"processed_files": 0, "folder": chat_id, "errors": ["No input text provided"]},
            "source_files": []
        }
    
    summarization_template = """    Tu es un assistant chargé d'analyser des informations textuelles sur des projets et de les formater dans un JSON spécifique.

    Voici les données textuelles à analyser:
    {text_data}

    Ta tâche est d'extraire des informations sur les projets mentionnés, y compris:
    1. Les noms des projets
    2. Un résumé des informations principales pour chaque projet
    3. Les avancements significatifs (points positifs)
    4. Les alertes mineures (points à surveiller)
    5. Les alertes critiques (problèmes urgents)
    6. Les événements à venir pour chaque projet ou catégorie (UNIQUEMENT s'ils sont explicitement mentionnés)

    IMPORTANT: Inclus TOUTES les informations dans le champ "information" de chaque projet. MAIS quand tu identifies une information comme étant un avancement, une alerte mineure ou une alerte critique, COPIE-LA ÉGALEMENT dans la catégorie correspondante (critical, small, advancements) pour qu'elle puisse être colorée. Ainsi, le texte apparaîtra dans le champ information mais sera automatiquement coloré.
    
    CONCERNANT LES ÉVÉNEMENTS À VENIR: Ne place des informations dans la section "upcoming_events" QUE s'il y a une mention EXPLICITE d'événements futurs, comme des phrases contenant "événements à venir", "semaine prochaine", "prochainement", etc. Si aucun événement futur n'est clairement mentionné, laisse la section "upcoming_events" VIDE avec un objet vide {{}}.
    
    Organise les informations selon le format JSON suivant:
    ```json
    {{
    "projects":{{
        "project1":{{
            "information":"",
            "critical":[],
            "small":[],
            "advancements":[]
        }},
        "project2":{{
            "subproject1":{{
                "information":"",
                "critical":[],
                "small":[],
                "advancements":[]
            }},
            "subproject2":{{
                "subsubproject1":{{
                    "information":"",
                    "critical":[],
                    "small":[],
                    "advancements":[]
                }},
                "subsubproject2":{{
                    "information":"",
                    "critical":[],
                    "small":[],
                    "advancements":[]
                }}
            }}
        }}
    }},
    "upcoming_events":{{
        "service1":[],
        "service2":[]
    }},
    "metadata":{{
        "processed_files": 1,
        "folder":"{chat_id_placeholder}" 
    }},
    "source_files":[
        {{
            "filename":"generated_from_text",
            "service_name":"Text Generator",
            "processed":true,
            "events_count":0
        }}
    ]
}}
    ```

    Assure-toi de:
    1. Identifier correctement les différents projets mentionnés dans le texte
    2. Créer un résumé concis et informatif pour chaque projet mais ne perdez pas de points importants
    3. Inclure TOUT le texte dans le champ "information", rien ne doit être perdu
    4. Ajouter AUSSI les informations importantes dans les catégories "advancements", "small", ou "critical" pour qu'elles soient colorées
    5. Organiser les événements à venir par catégories pertinentes UNIQUEMENT s'ils sont explicitement mentionnés
    6. Si aucun événement futur n'est mentionné (avec des termes comme "événements à venir", "semaine prochaine", etc.), LAISSER "upcoming_events" VIDE ({{}})
    7. Répondre UNIQUEMENT avec le JSON formaté, sans texte d'introduction ni d'explication
    8. Assurer que tout soit en Français
    9. Ne pas inventer de nouvelles informations, uniquement celles qui sont déjà présentes dans le texte
    10. Si aucun projet spécifique n'est identifiable, crée au moins un projet "Général" avec les informations disponibles
    11. Si tu n'as pas d'information sur les projets n'ajoute rien dans le JSON
    12. Remplace {chat_id_placeholder} par la valeur réelle de chat_id: {chat_id_value}
    """
    
    prompt = summarization_template.format(text_data=info, chat_id_placeholder=chat_id, chat_id_value=chat_id)
    
    try:
        prompt_size = len(prompt.encode('utf-8'))
        print(f"Generate PPTX from text prompt size: {prompt_size} bytes for chat {chat_id}")
        if prompt_size > 120000: 
            print(f"WARNING: Large prompt detected ({prompt_size} bytes) for chat {chat_id} for text generation, LLM may timeout/fail")
        
        print(f"Calling LLM for PPTX generation from text for chat {chat_id}...")
        time.sleep(1) 
        
        llm_response = summarize_model.invoke(prompt)
        print(f"LLM response received successfully for PPTX generation from text for chat {chat_id}")
        
        llm_response_cleaned = remove_tags_no_keep(llm_response, "<think>", "</think>")
        json_match = re.search(r'```json\\s*(.*?)```', llm_response_cleaned, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = llm_response_cleaned
        
        json_str = json_str.strip()
        result = json.loads(json_str)
        
        print(f"LLM PPTX generation from text completed successfully for chat {chat_id}")
        # Ensure metadata is consistent
        if "metadata" not in result: result["metadata"] = {}
        result["metadata"]["folder"] = chat_id
        result["metadata"].setdefault("processed_files", 1)
        if "source_files" not in result: result["source_files"] = [{"filename":"generated_from_text", "service_name":"Text Generator", "processed":True, "events_count":0}]

        return result
        
    except json.JSONDecodeError as json_e:
        print(f"JSON Decode Error during LLM text generation for chat {chat_id}: {str(json_e)}")
        print(f"LLM Response (cleaned) causing text gen error: '{json_str[:500]}...'")
        return {
            "projects": {"Erreur JSON": {"information": f"Erreur de décodage JSON: {str(json_e)}. Input: {info[:200]}", "critical": ["Erreur JSON"], "small": [], "advancements": []}},
            "upcoming_events": {},
            "metadata": {"processed_files": 1, "folder": chat_id, "error": f"LLM JSON Decode Error: {str(json_e)}"},
            "source_files": [{"filename": "generated_from_text_with_json_error", "processed": False, "error": str(json_e)}]
        }
    except Exception as e:
        error_str = str(e)
        print(f"Error during LLM PPTX generation from text for chat {chat_id}: {error_str}", exc_info=True)
        if any(keyword in error_str.lower() for keyword in ["eof", "connection", "timeout", "timed out"]):
            print("Connection/timeout error detected - likely Ollama service issue or timeout")
        elif "ollama" in error_str.lower():
            print("Ollama service error detected")
        
        print(f"Returning basic structure as fallback for chat {chat_id} due to LLM text generation error")
        return {
            "projects": {
                "Erreur de génération": {
                    "information": f"Une erreur s'est produite lors de la génération automatique à partir du texte: {error_str}. Contenu original: {info[:500]}...",
                    "critical": ["Erreur de génération LLM à partir du texte"],
                    "small": [],
                    "advancements": []
                }
            },
            "upcoming_events": {},
            "metadata": {"processed_files": 1, "folder": chat_id, "error": error_str},
            "source_files": [{"filename": "generated_from_text_with_error", "service_name": "Text Generator", "processed": False, "error": error_str}]
        }

if __name__ == "__main__":
    test_chat_id = "test_chat_extract_summarize"
    dummy_chat_folder = os.path.join(UPLOAD_FOLDER, test_chat_id)
    os.makedirs(dummy_chat_folder, exist_ok=True)
    
    # Create a dummy pptx file for testing the file processing path
    # You would need the python-pptx library to create a real one: pip install python-pptx
    # from pptx import Presentation
    # prs = Presentation()
    # prs.slides.add_slide(prs.slide_layouts[5])
    # prs.save(os.path.join(dummy_chat_folder, f"test_file_for_{test_chat_id}.pptx"))

    print(f"--- Testing aggregate_and_summarize with chat_id: {test_chat_id} ---")
    print("\n--- Test Case 1: No raw_structure_data (will attempt to process files if any in dummy folder) ---")
    result1 = aggregate_and_summarize(chat_id=test_chat_id, add_info="Test summary 1 for file processing")
    print("Result 1 (summarized from files/empty):")
    print(json.dumps(result1, indent=2, ensure_ascii=False))

    print("\n--- Test Case 2: With raw_structure_data ---")
    dummy_structure = {
        "projects": {"CachedProject": {"information": "This is cached info.", "critical": ["Cached critical alert"], "small": [], "advancements": ["Cached advancement"]}},
        "upcoming_events": {"CachedService": ["Cached upcoming event"]},
        "metadata": {"processed_files": 1, "folder": test_chat_id, "errors": [], "source_files": [{"filename": "from_cache.pptx", "service_name":"CacheServ", "processed":True}]}, # Added source_files to dummy
    }
    result2 = aggregate_and_summarize(chat_id=test_chat_id, add_info="Test summary 2 using cache", raw_structure_data=dummy_structure)
    print("Result 2 (summarized from raw_structure_data):")
    print(json.dumps(result2, indent=2, ensure_ascii=False))

    print("\n--- Testing Generate_pptx_from_text ---")
    text_info_for_generation = "Le projet Phoenix est en bonne voie. Une alerte mineure concerne le budget. Prochain jalon: livraison beta la semaine prochaine."
    result3 = Generate_pptx_from_text(chat_id=test_chat_id, info=text_info_for_generation)
    print("Result 3 (generated from text):")
    print(json.dumps(result3, indent=2, ensure_ascii=False))

    # import shutil
    # if os.path.exists(dummy_chat_folder):
    #     shutil.rmtree(dummy_chat_folder)
    #     print(f"\nCleaned up dummy folder: {dummy_chat_folder}")
