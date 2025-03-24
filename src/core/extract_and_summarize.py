import os,sys
import re
from pptx import Presentation

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
    Returns a dictionary with common_info, upcoming_info, advancements, small_alerts, and critical_alerts
    for updating the PowerPoint template.
    
    Parameters:
      pptx_folder (str): Path to the folder containing PowerPoint files to analyze.
    """
    project_data = {}
    file_count = 0
    
    # Check if the folder exists
    if not os.path.exists(pptx_folder):
        print(f"Warning: Folder {pptx_folder} does not exist.")
        return {
            "common_info": "Aucune information disponible - dossier non trouvé.",
            "upcoming_info": "Aucun événement particulier prévu.",
            "advancements": "Aucun avancement significatif à signaler.",
            "small_alerts": "Aucune alerte mineure à signaler.",
            "critical_alerts": "Aucune alerte critique à signaler."
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
    
    # If no files were processed, return empty data
    if file_count == 0:
        return {
            "common_info": "Aucune information disponible - aucun fichier PPTX trouvé.",
            "upcoming_info": "Aucun événement particulier prévu.",
            "advancements": "Aucun avancement significatif à signaler.",
            "small_alerts": "Aucune alerte mineure à signaler.",
            "critical_alerts": "Aucune alerte critique à signaler."
        }
    
    # Extract common information, upcoming work information, and alerts
    extracted_data = extract_common_and_upcoming_info(project_data)
    
    # Summarize common information if it's too long
    if len(extracted_data["common_info"]) > 2000:  # Arbitrary threshold, adjust as needed
        prompt = (
            "Voici des informations agrégées à partir de plusieurs fichiers PowerPoint. "
            "Votre tâche est de créer un résumé structuré en français qui capture les points clés "
            "et les idées principales de manière concise et informative.\n\n"
            f"{extracted_data['common_info']}"
        )
        extracted_data["common_info"] = summarize_model.invoke(prompt).content
        extracted_data["common_info"] = remove_tags_no_keep(extracted_data["common_info"], "<think>", "</think>")
    
    # Summarize upcoming information if it's too long
    if len(extracted_data["upcoming_info"]) > 1000:  # Arbitrary threshold, adjust as needed
        prompt = (
            "Voici des informations concernant les événements de la semaine à venir. "
            "Veuillez résumer ces informations de manière concise en français, en conservant "
            "les points essentiels concernant les prochaines étapes et événements.\n\n"
            f"{extracted_data['upcoming_info']}"
        )
        extracted_data["upcoming_info"] = summarize_model.invoke(prompt).content
        extracted_data["upcoming_info"] = remove_tags_no_keep(extracted_data["upcoming_info"], "<think>", "</think>")
    
    # Summarize advancements if there are many
    if len(extracted_data["advancements"]) > 1000:
        prompt = (
            "Voici des informations concernant les avancements de différents projets. "
            "Veuillez résumer ces avancements de manière concise en français, en mettant en avant "
            "les progrès les plus significatifs.\n\n"
            f"{extracted_data['advancements']}"
        )
        extracted_data["advancements"] = summarize_model.invoke(prompt).content
        extracted_data["advancements"] = remove_tags_no_keep(extracted_data["advancements"], "<think>", "</think>")
    
    # Summarize small alerts if there are many
    if len(extracted_data["small_alerts"]) > 1000:
        prompt = (
            "Voici des informations concernant les alertes mineures de différents projets. "
            "Veuillez résumer ces alertes de manière concise en français, en conservant "
            "les informations essentielles sur les points à surveiller.\n\n"
            f"{extracted_data['small_alerts']}"
        )
        extracted_data["small_alerts"] = summarize_model.invoke(prompt).content
        extracted_data["small_alerts"] = remove_tags_no_keep(extracted_data["small_alerts"], "<think>", "</think>")
    
    # Summarize critical alerts if there are many
    if len(extracted_data["critical_alerts"]) > 1000:
        prompt = (
            "Voici des informations concernant les alertes critiques de différents projets. "
            "Veuillez résumer ces alertes de manière concise en français, en mettant en évidence "
            "les problèmes les plus urgents qui nécessitent une attention immédiate.\n\n"
            f"{extracted_data['critical_alerts']}"
        )
        extracted_data["critical_alerts"] = summarize_model.invoke(prompt).content
        extracted_data["critical_alerts"] = remove_tags_no_keep(extracted_data["critical_alerts"], "<think>", "</think>")
    
    return extracted_data

if __name__ == "__main__":
    folder = "pptx_folder"  # Update with your actual folder path
    result = aggregate_and_summarize(folder)
    print("Common Info:", result["common_info"])
    print("Upcoming Info:", result["upcoming_info"])
    print("Advancements:", result["advancements"])
    print("Small Alerts:", result["small_alerts"])
    print("Critical Alerts:", result["critical_alerts"])
