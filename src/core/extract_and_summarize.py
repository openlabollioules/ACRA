import os,sys
import re
import json
from pptx import Presentation
from langchain_core.prompts import PromptTemplate
from copy import deepcopy

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from langchain_ollama import OllamaLLM
summarize_model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", temperature=0.7, num_ctx=64000)

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
      dict: A nested dictionary with project/subproject structure containing information and alerts
    """
    # New project structure
    aggregated_data = {}
    file_count = 0
    
    # Check if the folder exists
    if not os.path.exists(pptx_folder):
        print(f"Warning: Folder {pptx_folder} does not exist.")
        return {}
    
    # Get all PPTX files in the folder
    for filename in os.listdir(pptx_folder):
        if filename.lower().endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            # Extract project data from the presentation
            try:
                file_project_data = extract_projects_from_presentation(file_path)
                file_count += 1
                
                # Process projects data from file_project_data
                if "projects" in file_project_data:
                    for main_project, subprojects in file_project_data["projects"].items():
                        # Ensure the main project exists in aggregated data
                        if main_project not in aggregated_data:
                            aggregated_data[main_project] = {}
                        
                        # Process each subproject
                        for subproject_name, subproject_info in subprojects.items():
                            # If the subproject already exists, merge information
                            if subproject_name in aggregated_data[main_project]:
                                # Merge information text
                                if "information" in subproject_info and subproject_info["information"]:
                                    existing_info = aggregated_data[main_project][subproject_name].get("information", "")
                                    aggregated_data[main_project][subproject_name]["information"] = (
                                        existing_info + "\n" + subproject_info["information"] if existing_info else subproject_info["information"]
                                    )
                                
                                # Merge alerts
                                for alert_type in ["critical", "small", "advancements"]:
                                    if alert_type in subproject_info and subproject_info[alert_type]:
                                        aggregated_data[main_project][subproject_name][alert_type].extend(
                                            subproject_info[alert_type]
                                        )
                                
                                # Merge upcoming events
                                if "upcoming_events" in subproject_info and subproject_info["upcoming_events"]:
                                    aggregated_data[main_project][subproject_name]["upcoming_events"].extend(
                                        [event for event in subproject_info["upcoming_events"] 
                                         if event not in aggregated_data[main_project][subproject_name]["upcoming_events"]]
                                    )
                            else:
                                # Add new subproject
                                aggregated_data[main_project][subproject_name] = deepcopy(subproject_info)
            except Exception as e:
                print(f"Error processing file {filename}: {str(e)}")
    
    # If no files were processed, return empty data structure
    if file_count == 0:
        return {}
    
    # Prepare the data to send to the LLM for summarization if needed
    prompt_inputs = {
        "project_data": json.dumps(aggregated_data, indent=2, ensure_ascii=False)
    }
    
    # Create a prompt template for the LLM
    summarization_template = PromptTemplate.from_template("""
    Tu es un assistant chargé de résumer des informations de projets et de les formater.

    Voici les données des projets:
    {project_data}
    
    Analyse ces données et identifie les points clés pour chaque projet et sous-projet.
    Pour chaque entrée, tu peux conserver la structure mais synthétise les informations
    pour qu'elles soient plus concises tout en préservant les détails importants.
    
    Les alertes critiques, alertes mineures et avancements doivent être conservés tels quels,
    mais tu peux éliminer les redondances éventuelles.
    
    Réponds uniquement avec la structure JSON modifiée, sans texte d'introduction ni d'explication.
    """)
    
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
        
        return result
        
    except Exception as e:
        print(f"Error during LLM summarization: {str(e)}")
        # If summarization fails, return the raw aggregated data
        return aggregated_data

if __name__ == "__main__":
    folder = "pptx_folder"  # Update with your actual folder path
    result = aggregate_and_summarize(folder)
    print("Activities:", result["activities"])
    print("Upcoming Events:", result["upcoming_events"])
