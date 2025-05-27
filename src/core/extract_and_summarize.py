import os,sys
import re
import json
from langchain_core.prompts import PromptTemplate
from dotenv import load_dotenv
import time

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Load environment variables
load_dotenv()
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")

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

def aggregate_and_summarize(pptx_folder, add_info=None):
    """
    Main function to aggregate the IF texts from all PPTX files in the folder and obtain a summarized result.
    Uses an LLM to summarize the project information and return it in the specified JSON format.
    
    Parameters:
      pptx_folder (str): Path to the folder containing PowerPoint files to analyze.
      add_info (str, optional): Additional information to include in the summary.
    Returns:
      dict: A nested dictionary with project/subproject structure containing information and alerts
    """
    # New project structure
    aggregated_data = {}
    file_count = 0
    processed_files = []
    extraction_errors = []
    
    # Determine the full path to the folder
    full_path = pptx_folder
    if not os.path.isabs(pptx_folder):
        full_path = os.path.join(UPLOAD_FOLDER, pptx_folder)
    
    print(f"Processing folder: {full_path}")
    
    # Check if the folder exists
    if not os.path.exists(full_path):
        error_msg = f"Warning: Folder {full_path} does not exist."
        print(error_msg)
        extraction_errors.append(error_msg)
        return {
            "projects": {},
            "upcoming_events": {},
            "metadata": {
                "processed_files": 0,
                "folder": os.path.basename(pptx_folder),
                "errors": extraction_errors
            },
            "source_files": []
        }
    
    # List all files in the folder to diagnose issues
    all_files = os.listdir(full_path)
    print(f"Files in folder: {all_files}")
    pptx_files = [f for f in all_files if f.lower().endswith(".pptx")]
    print(f"PPTX files found: {pptx_files}")
    
    if not pptx_files:
        error_msg = f"No PowerPoint files found in folder {full_path}"
        print(error_msg)
        extraction_errors.append(error_msg)
        return {
            "projects": {},
            "upcoming_events": {},
            "metadata": {
                "processed_files": 0,
                "folder": os.path.basename(pptx_folder),
                "errors": extraction_errors
            },
            "source_files": []
        }
    
    # Get all PPTX files in the folder
    for filename in pptx_files:
        file_path = os.path.join(full_path, filename)
        print(f"Processing file: {file_path}")
        
        # Extract project data from the presentation
        try:
            file_project_data = extract_projects_from_presentation(file_path)
            file_count += 1
            
            # Add processed file info
            service_name = os.path.basename(file_path).split('_')[-1].replace('.pptx', '')
            processed_file_info = {
                "filename": filename,
                "service_name": service_name,
                "processed": True
            }
            
            # Check if any projects were extracted
            if "projects" in file_project_data and file_project_data["projects"]:
                print(f"Successfully extracted projects from {filename}")
                project_count = len(file_project_data["projects"])
                processed_file_info["project_count"] = project_count
            else:
                # Check for error message in metadata
                if "metadata" in file_project_data and "error" in file_project_data["metadata"]:
                    error = file_project_data["metadata"]["error"]
                    print(f"Error in file {filename}: {error}")
                    extraction_errors.append(f"File {filename}: {error}")
                    processed_file_info["error"] = error
                else:
                    warning = f"No projects extracted from {filename}"
                    print(warning)
                    extraction_errors.append(warning)
                    processed_file_info["warning"] = warning
            
            processed_files.append(processed_file_info)
            
            # Process projects data from file_project_data
            if "projects" in file_project_data:
                for main_project, main_project_content in file_project_data["projects"].items():
                    # Ensure the main project exists in aggregated data
                    if main_project not in aggregated_data:
                        aggregated_data[main_project] = {}
                    
                    # Check if main_project_content is a terminal node or contains subprojects
                    is_terminal = "information" in main_project_content
                    
                    if is_terminal:
                        # This is a terminal node, merge the data directly
                        if "information" in aggregated_data[main_project]:
                            # Merge with existing data
                            aggregated_data[main_project]["information"] += "\n" + main_project_content["information"] if aggregated_data[main_project]["information"] else main_project_content["information"]
                            
                            # Merge alerts and advancements
                            for alert_type in ["critical", "small", "advancements"]:
                                if alert_type in main_project_content:
                                    if alert_type not in aggregated_data[main_project]:
                                        aggregated_data[main_project][alert_type] = []
                                    aggregated_data[main_project][alert_type].extend(
                                        item for item in main_project_content[alert_type] 
                                        if item not in aggregated_data[main_project][alert_type]
                                    )
                        else:
                            # Copy the data for a new terminal node
                            aggregated_data[main_project] = {
                                "information": main_project_content.get("information", ""),
                                "critical": main_project_content.get("critical", []),
                                "small": main_project_content.get("small", []),
                                "advancements": main_project_content.get("advancements", [])
                            }
                    else:
                        # This contains subprojects
                        for subproject_name, subproject_content in main_project_content.items():
                            # Skip metadata fields that might be in the dictionary
                            if subproject_name in ["information", "critical", "small", "advancements"]:
                                # Handle top-level project information if it exists alongside subprojects
                                if subproject_name == "information" and subproject_content:
                                    if "information" not in aggregated_data[main_project]:
                                        aggregated_data[main_project]["information"] = subproject_content
                                    else:
                                        aggregated_data[main_project]["information"] += "\n" + subproject_content
                                elif subproject_name in ["critical", "small", "advancements"] and subproject_content:
                                    if subproject_name not in aggregated_data[main_project]:
                                        aggregated_data[main_project][subproject_name] = []
                                    aggregated_data[main_project][subproject_name].extend(
                                        item for item in subproject_content 
                                        if item not in aggregated_data[main_project][subproject_name]
                                    )
                                continue
                            
                            # Process the subproject
                            if subproject_name not in aggregated_data[main_project]:
                                aggregated_data[main_project][subproject_name] = {}
                            
                            # Check if subproject_content is a terminal node or further nested
                            sub_is_terminal = "information" in subproject_content
                            
                            if sub_is_terminal:
                                # This is a terminal subproject
                                if "information" in aggregated_data[main_project][subproject_name]:
                                    # Merge with existing data
                                    aggregated_data[main_project][subproject_name]["information"] += "\n" + subproject_content["information"] if aggregated_data[main_project][subproject_name]["information"] else subproject_content["information"]
                                    
                                    # Merge alerts and advancements
                                    for alert_type in ["critical", "small", "advancements"]:
                                        if alert_type in subproject_content:
                                            if alert_type not in aggregated_data[main_project][subproject_name]:
                                                aggregated_data[main_project][subproject_name][alert_type] = []
                                            aggregated_data[main_project][subproject_name][alert_type].extend(
                                                item for item in subproject_content[alert_type] 
                                                if item not in aggregated_data[main_project][subproject_name][alert_type]
                                            )
                                else:
                                    # Copy the data for a new terminal subproject
                                    aggregated_data[main_project][subproject_name] = {
                                        "information": subproject_content.get("information", ""),
                                        "critical": subproject_content.get("critical", []),
                                        "small": subproject_content.get("small", []),
                                        "advancements": subproject_content.get("advancements", [])
                                    }
                            else:
                                # This contains sub-subprojects
                                for subsubproject_name, subsubproject_content in subproject_content.items():
                                    # Skip metadata fields
                                    if subsubproject_name in ["information", "critical", "small", "advancements"]:
                                        # Handle subproject information if it exists alongside sub-subprojects
                                        if subsubproject_name == "information" and subsubproject_content:
                                            if "information" not in aggregated_data[main_project][subproject_name]:
                                                aggregated_data[main_project][subproject_name]["information"] = subsubproject_content
                                            else:
                                                aggregated_data[main_project][subproject_name]["information"] += "\n" + subsubproject_content
                                        elif subsubproject_name in ["critical", "small", "advancements"] and subsubproject_content:
                                            if subsubproject_name not in aggregated_data[main_project][subproject_name]:
                                                aggregated_data[main_project][subproject_name][subsubproject_name] = []
                                            aggregated_data[main_project][subproject_name][subsubproject_name].extend(
                                                item for item in subsubproject_content 
                                                if item not in aggregated_data[main_project][subproject_name][subsubproject_name]
                                            )
                                        continue
                                    
                                    # Process the sub-subproject (assuming it's always a terminal node)
                                    if subsubproject_name not in aggregated_data[main_project][subproject_name]:
                                        aggregated_data[main_project][subproject_name][subsubproject_name] = {
                                            "information": subsubproject_content.get("information", ""),
                                            "critical": subsubproject_content.get("critical", []),
                                            "small": subsubproject_content.get("small", []),
                                            "advancements": subsubproject_content.get("advancements", [])
                                        }
                                    else:
                                        # Merge with existing data
                                        if "information" in subsubproject_content:
                                            if "information" in aggregated_data[main_project][subproject_name][subsubproject_name]:
                                                aggregated_data[main_project][subproject_name][subsubproject_name]["information"] += "\n" + subsubproject_content["information"]
                                            else:
                                                aggregated_data[main_project][subproject_name][subsubproject_name]["information"] = subsubproject_content["information"]
                                        
                                        # Merge alerts and advancements
                                        for alert_type in ["critical", "small", "advancements"]:
                                            if alert_type in subsubproject_content:
                                                if alert_type not in aggregated_data[main_project][subproject_name][subsubproject_name]:
                                                    aggregated_data[main_project][subproject_name][subsubproject_name][alert_type] = []
                                                aggregated_data[main_project][subproject_name][subsubproject_name][alert_type].extend(
                                                    item for item in subsubproject_content[alert_type] 
                                                    if item not in aggregated_data[main_project][subproject_name][subsubproject_name][alert_type]
                                                )
            
            # Handle upcoming events from metadata
            if "metadata" in file_project_data and "collected_upcoming_events" in file_project_data["metadata"]:
                # Need to know the service name to categorize events
                service_name = os.path.basename(file_path).split('_')[-1].replace('.pptx', '')
                events = file_project_data["metadata"]["collected_upcoming_events"]
                
                if events:
                    processed_file_info["events_count"] = len(events)
                    
                    if "upcoming_events" not in aggregated_data:
                        aggregated_data["upcoming_events"] = {}
                    
                    if service_name not in aggregated_data["upcoming_events"]:
                        aggregated_data["upcoming_events"][service_name] = []
                    
                    for event in events:
                        if event not in aggregated_data["upcoming_events"][service_name]:
                            aggregated_data["upcoming_events"][service_name].append(event)
                else:
                    processed_file_info["events_count"] = 0
        except Exception as e:
            error_message = f"Error processing file {filename}: {str(e)}"
            print(error_message)
            extraction_errors.append(error_message)
            processed_files.append({
                "filename": filename, 
                "processed": False,
                "error": str(e)
            })
    
    # If no projects were extracted but files were processed, that's a problem
    if file_count > 0 and not aggregated_data:
        print(f"WARNING: {file_count} files were processed but no project data was extracted")
        extraction_errors.append(f"{file_count} files were processed but no project data was extracted")
    
    # If no files were processed, return empty data structure with error info
    if file_count == 0:
        error_msg = "No files were successfully processed"
        print(error_msg)
        extraction_errors.append(error_msg)
        return {
            "projects": {},
            "upcoming_events": {},
            "metadata": {
                "processed_files": 0,
                "folder": os.path.basename(pptx_folder),
                "errors": extraction_errors
            },
            "source_files": processed_files
        }
    
    # Prepare the data structure for return
    result = {
        "projects": aggregated_data.get("projects", aggregated_data),
        "upcoming_events": aggregated_data.get("upcoming_events", {}),
        "metadata": {
            "processed_files": file_count,
            "folder": os.path.basename(pptx_folder),
            "errors": extraction_errors if extraction_errors else []
        },
        "source_files": processed_files
    }
    
    # Remove upcoming_events from projects if it was accidentally included there
    if "upcoming_events" in result["projects"]:
        del result["projects"]["upcoming_events"]
    
    # Log the size and structure of the result
    print(f"Final result structure: {len(result['projects'])} projects, {len(result['upcoming_events'])} services with events")
    if not result["projects"]:
        print("WARNING: No projects were extracted from any files")
    
    # Prepare the data to send to the LLM for summarization if needed
    prompt_inputs = {
        "project_data": json.dumps(result, indent=2, ensure_ascii=False),
        "temp_add_info": ""
    }
    if add_info:
        prompt_inputs["temp_add_info"] = f"Voici des informations supplémentaires qui peuvent être utiles pour la synthèse: {add_info}"
    
    # Create a prompt template for the LLM
    summarization_template = """
    Tu es un assistant chargé de résumer des informations de projets et de les formater.

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
    
    # Generate the prompt using direct format method
    prompt = summarization_template.format(**prompt_inputs)
    
    # Only try to use LLM if we have actual project data
    if not result["projects"]:
        print("Skipping LLM summarization because no project data was extracted")
        return result
    
    # Call the LLM to generate the summary in JSON format
    try:
        # Check prompt size to warn about potential issues
        prompt_size = len(prompt.encode('utf-8'))
        print(f"Summarization prompt size: {prompt_size} bytes")
        if prompt_size > 100000:  # 100KB threshold
            print(f"WARNING: Large prompt detected ({prompt_size} bytes), LLM may timeout or fail")
        
        print("Calling LLM for summarization...")
        llm_response = summarize_model.invoke(prompt)
        print("LLM response received successfully")
        
        # Extract the JSON part from the response
        llm_response = remove_tags_no_keep(llm_response, "<think>", "</think>")
        json_match = re.search(r'```json\s*(.*?)```', llm_response, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = llm_response
        
        # Clean the JSON string and parse it
        json_str = json_str.strip()
        summarized_result = json.loads(json_str)
        
        print("LLM summarization completed successfully")
        return summarized_result
        
    except Exception as e:
        print(f"Error during LLM summarization: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        
        # Check if it's a connection-related error
        if "EOF" in str(e) or "Connection" in str(e) or "timeout" in str(e).lower():
            print("Connection error detected - likely Ollama service issue or timeout")
        
        # If summarization fails, return the raw aggregated data
        print("Returning raw aggregated data as fallback")
        return result

def Generate_pptx_from_text(pptx_folder, info=None):
    """
    Generate a JSON structure from text input that can be used by update_table_with_project_data.
    Uses an LLM to process the text information and return it in the specified JSON format.
    
    Parameters:
      pptx_folder (str): Path to the folder (used for compatibility, not used in processing).
      info (str): Text information to process and structure into JSON format.
    
    Returns:
      dict: A dictionary with activities containing project information and upcoming_events in the specified JSON format
    """
    # If no info is provided, return empty structure
    if not info:
        return {
            "projects": {},
            "upcoming_events": {},
            "metadata": {
                "processed_files": 0,
                "folder": os.path.basename(pptx_folder)
            },
            "source_files": []
        }
    
    # Create a prompt template for the LLM
    summarization_template = """
    Tu es un assistant chargé d'analyser des informations textuelles sur des projets et de les formater dans un JSON spécifique.

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
        "folder":"Name of the folder"
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
    """
    
    # Format the prompt directly with f-string instead of using PromptTemplate
    prompt = summarization_template.format(text_data=info)
    
    # Call the LLM to generate the summary in JSON format
    try:
        # Check prompt size to warn about potential issues
        prompt_size = len(prompt.encode('utf-8'))
        print(f"Generate PPTX prompt size: {prompt_size} bytes")
        if prompt_size > 100000:  # 100KB threshold
            print(f"WARNING: Large prompt detected ({prompt_size} bytes), LLM may timeout or fail")
        
        print("Calling LLM for PPTX generation...")
        
        # Add a small delay to prevent immediate consecutive LLM calls
        time.sleep(1)  # 1 second delay
        
        llm_response = summarize_model.invoke(prompt)
        print("LLM response received successfully")
        
        # Extract the JSON part from the response
        llm_response = remove_tags_no_keep(llm_response, "<think>", "</think>")
        json_match = re.search(r'```json\s*(.*?)```', llm_response, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = llm_response
        
        # Clean the JSON string and parse it
        json_str = json_str.strip()
        result = json.loads(json_str)
        print("result :", result)
        print("result type :", type(result))
        
        print("LLM PPTX generation completed successfully")
        return result
        
    except Exception as e:
        error_str = str(e)
        print(f"Error during LLM PPTX generation: {error_str}")
        print(f"Error type: {type(e).__name__}")
        
        # Check if it's a connection-related error
        if any(keyword in error_str.lower() for keyword in ["eof", "connection", "timeout", "timed out"]):
            print("Connection/timeout error detected during PPTX generation - likely Ollama service issue or timeout")
        elif "ollama" in error_str.lower():
            print("Ollama service error detected during PPTX generation")
        
        # If generation fails, return a basic structure with error info
        print("Returning basic structure as fallback due to LLM error")
        return {
            "projects": {
                "Erreur de génération": {
                    "information": f"Une erreur s'est produite lors de la génération automatique: {error_str}. Contenu original: {info[:500]}...",
                    "critical": ["Erreur de génération LLM"],
                    "small": [],
                    "advancements": []
                }
            },
            "upcoming_events": {},
            "metadata": {
                "processed_files": 1,
                "folder": os.path.basename(pptx_folder),
                "error": error_str
            },
            "source_files": [
                {
                    "filename": "generated_from_text_with_error",
                    "service_name": "Text Generator",
                    "processed": False,
                    "error": error_str
                }
            ]
        }

if __name__ == "__main__":
    folder = "pptx_folder"  # Update with your actual folder path
    result = aggregate_and_summarize(folder)
    print("Activities:", result["activities"])
    print("Upcoming Events:", result["upcoming_events"])
