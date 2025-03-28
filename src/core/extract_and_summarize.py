import os,sys
import re
import json
from pptx import Presentation
from langchain_core.prompts import PromptTemplate

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from config import summarize_model
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

def aggregate_and_summarize(pptx_folder):
    """
    Main function to aggregate the IF texts from all PPTX files in the folder and obtain a summarized result.
    Uses an LLM to summarize the project information and return it in the specified JSON format.
    
    Parameters:
      pptx_folder (str): Path to the folder containing PowerPoint files to analyze.
    
    Returns:
      dict: A dictionary with activities containing project information and upcoming_events in the specified JSON format
    """
    project_data = {}
    file_count = 0
    
    # Check if the folder exists
    if not os.path.exists(pptx_folder):
        print(f"Warning: Folder {pptx_folder} does not exist.")
        return {
            "activities": {},
            "upcoming_events": {
                "General": "Aucun événement particulier prévu."
            }
        }
    
    # Get all PPTX files in the folder
    for filename in os.listdir(pptx_folder):
        if filename.lower().endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            # Extract project data from the presentation
            try:
                file_project_data = extract_projects_from_presentation(file_path)
                file_count += 1
                
                # Process activities from file_project_data
                if "activities" in file_project_data:
                    for project_name, project_info in file_project_data["activities"].items():
                        # If the project already exists, merge the information
                        if project_name in project_data:
                            # Merge information
                            if "information" in project_info and project_info["information"]:
                                if "information" not in project_data[project_name]:
                                    project_data[project_name]["information"] = project_info["information"]
                                else:
                                    project_data[project_name]["information"] += "\n" + project_info["information"]
                            
                            # Merge alerts
                            if "alerts" in project_info:
                                if "alerts" not in project_data[project_name]:
                                    project_data[project_name]["alerts"] = {
                                        "advancements": [],
                                        "small_alerts": [],
                                        "critical_alerts": []
                                    }
                                
                                # Merge advancements
                                if "advancements" in project_info["alerts"]:
                                    project_data[project_name]["alerts"]["advancements"].extend(
                                        project_info["alerts"]["advancements"]
                                    )
                                
                                # Merge small alerts
                                if "small_alerts" in project_info["alerts"]:
                                    project_data[project_name]["alerts"]["small_alerts"].extend(
                                        project_info["alerts"]["small_alerts"]
                                    )
                                
                                # Merge critical alerts
                                if "critical_alerts" in project_info["alerts"]:
                                    project_data[project_name]["alerts"]["critical_alerts"].extend(
                                        project_info["alerts"]["critical_alerts"]
                                    )
                        else:
                            # Add new project
                            project_data[project_name] = project_info
                
                # Add upcoming_events if available
                if "upcoming_events" in file_project_data:
                    if "upcoming_events" not in project_data:
                        project_data["upcoming_events"] = file_project_data["upcoming_events"]
                    else:
                        project_data["upcoming_events"] += "\n" + file_project_data["upcoming_events"]
            except Exception as e:
                print(f"Error processing file {filename}: {str(e)}")
    
    # If no files were processed, return empty data structure
    if file_count == 0:
        return {
            "activities": {},
            "upcoming_events": {
                "General": "Aucun événement particulier prévu."
            }
        }
    
    # Prepare the data to send to the LLM
    processed_data = {
        "projects": {},
        "events": project_data.get("upcoming_events", "")
    }
    
    # Extract essential information from each project
    for project_name, project_info in project_data.items():
        # Skip the "upcoming_events" key as it's not a project
        if project_name == "upcoming_events":
            continue
        
        processed_data["projects"][project_name] = {
            "information": project_info.get("information", "").strip(),
            "advancements": project_info.get("alerts", {}).get("advancements", []),
            "small_alerts": project_info.get("alerts", {}).get("small_alerts", []),
            "critical_alerts": project_info.get("alerts", {}).get("critical_alerts", [])
        }
    
    # Create a prompt template for the LLM
    summarization_template = PromptTemplate.from_template("""
    Tu es un assistant chargé de résumer des informations de projets et de les formater dans un JSON spécifique.

    Voici les données des projets:
    {project_data}
    
    Voici les événements à venir:
    {events_data}

    Crée un résumé concis pour chaque projet et organise les informations selon le format JSON suivant:
    ```json
    {{
      "activities": {{
        "Nom du Projet": {{
          "summary": "Résumé concis des informations principales du projet en une ou deux phrases",
          "alerts": {{
            "advancements": ["Liste des avancements significatifs, sous forme de points concis"],
            "small_alerts": ["Liste des alertes mineures, sous forme de points concis"],
            "critical_alerts": ["Liste des alertes critiques, sous forme de points concis"]
          }}
        }},
        // Autres projets...
      }},
      "upcoming_events": {{
        "Catégorie1": "Description des événements à venir pour cette catégorie",
        "Catégorie2": "Description des événements à venir pour cette catégorie",
        // Autres catégories...
      }}
    }}
    ```

    Assure-toi de:
    1. Créer un résumé synthétique et informatif pour chaque projet
    2. Conserver les informations essentielles des alertes (advancements, small_alerts, critical_alerts)
    3. Organiser les événements à venir par catégories pertinentes (ex: Cybersecurity, UX/UI Design, Consulting)
    4. Répondre UNIQUEMENT avec le JSON formaté, sans texte d'introduction ni d'explication
    5. Assure toi que tout soit en Français.

    Les alertes ne doivent contenir que les points vraiment importants, pas besoin de tout inclure.
    """)
    
    # Prepare the inputs for the prompt
    prompt_inputs = {
        "project_data": json.dumps(processed_data["projects"], indent=2, ensure_ascii=True),
        "events_data": processed_data["events"]
    }
    
    # Generate the prompt
    prompt = summarization_template.format(**prompt_inputs)
    
    # Call the LLM to generate the summary in JSON format
    try:
        llm_response = summarize_model.invoke(prompt)
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
        
        # Ensure the expected structure exists
        if "activities" not in result:
            result["activities"] = {}
        if "upcoming_events" not in result:
            result["upcoming_events"] = {"General": "Aucun événement particulier prévu."}
        
        return result
        
    except Exception as e:
        print(f"Error during LLM summarization: {str(e)}")
        # Create a basic structure with the raw data as fallback
        result = {
            "activities": {},
            "upcoming_events": {}
        }
        
        # Process upcoming events
        upcoming_texts = project_data.get("upcoming_events", "")
        if upcoming_texts:
            # Try to identify categories in the text
            categories = re.findall(r"([A-Za-z0-9/]+):\s*([^:]+?)(?=\n[A-Za-z0-9/]+:|$)", upcoming_texts, re.DOTALL)
            
            if categories:
                for category, text in categories:
                    result["upcoming_events"][category.strip()] = text.strip()
            else:
                result["upcoming_events"]["General"] = upcoming_texts.strip()
        else:
            result["upcoming_events"]["General"] = "Aucun événement particulier prévu."
        
        # Process projects without LLM summarization
        for project_name, project_info in project_data.items():
            if project_name == "upcoming_events":
                continue
                
            summary = project_info.get("information", "").strip()
            
            result["activities"][project_name] = {
                "summary": summary if summary else "Aucune information disponible.",
                "alerts": {
                    "advancements": project_info.get("alerts", {}).get("advancements", []),
                    "small_alerts": project_info.get("alerts", {}).get("small_alerts", []),
                    "critical_alerts": project_info.get("alerts", {}).get("critical_alerts", [])
                }
            }
        
        return result

if __name__ == "__main__":
    folder = "pptx_folder"  # Update with your actual folder path
    result = aggregate_and_summarize(folder)
    print("Activities:", result["activities"])
    print("Upcoming Events:", result["upcoming_events"])
