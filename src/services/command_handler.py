"""
Command Handler Service for ACRA
Centralized command processing for the pipeline
"""
import os
import json
import datetime
from typing import Dict, Any, Generator, List, Tuple
from OLLibrary.utils.log_service import get_logger
from OLLibrary.utils.json_service import extract_json
from config_pipeline import acra_config
from .file_manager import FileManager
from .model_manager import model_manager
from .cleanup_service import cleanup_orphaned_folders

log = get_logger(__name__)

class CommandHandler:
    """
    Centralized command handler for ACRA pipeline commands.
    Processes commands like /summarize, /structure, /clear, etc.
    """
    
    def __init__(self, file_manager: FileManager):
        self.file_manager = file_manager
        self.cached_structure = None
        self.waiting_for_confirmation = False
        self.confirmation_command = ""
        self.confirmation_additional_info = ""
        self.last_response = None
    
    def reset_state(self):
        """Reset command handler state"""
        self.cached_structure = None
        self.waiting_for_confirmation = False
        self.confirmation_command = ""
        self.confirmation_additional_info = ""
        self.last_response = None
    
    def get_available_commands(self) -> str:
        """Get list of available commands"""
        return """Les commandes sont les suivantes : 

/summarize [instructions] --> Affiche les résumés existants et demande confirmation avant d'en générer un nouveau. Vous pouvez ajouter des instructions spécifiques après la commande pour guider le résumé.
/structure --> Renvoie la structure des fichiers 
/clear [IDs] --> Nettoie tous les dossiers orphelins et supprime les fichiers associés dans OpenWebUI (préserve la conversation actuelle et éventuellement d'autres IDs spécifiés)
/generate --> Génère tout le pptx en fonction du texte ( /generate [Avancements de la semaine])
/merge --> Fusionne tous les fichiers pptx envoyés
/regroup --> Regroupe les informations des projets similaires ou liés"""
    
    def handle_confirmation(self, message: str) -> Tuple[bool, str]:
        """
        Handle confirmation responses (yes/no).
        
        Returns:
            Tuple[bool, str]: (handled, response_message)
        """
        if not self.waiting_for_confirmation:
            return False, ""
        
        message_lower = message.lower()
        
        if message_lower in ["yes", "y", "oui", "o"]:
            self.waiting_for_confirmation = False
            
            if self.confirmation_command == "summarize":
                return True, self._execute_summarize(self.confirmation_additional_info)
            
        elif message_lower in ["no", "n", "non"]:
            self.waiting_for_confirmation = False
            return True, "Génération de résumé annulée."
        
        # Reset if we get any other input
        self.waiting_for_confirmation = False
        return False, ""
    
    def handle_summarize_command(self, message: str) -> str:
        """Handle /summarize command"""
        # Extract additional information after the command
        additional_info = None
        if " " in message:
            command_parts = message.split(" ", 1)
            if len(command_parts) > 1 and command_parts[1].strip():
                additional_info = command_parts[1].strip()
        
        # Get existing summaries
        existing_summaries = self.file_manager.get_existing_summaries()
        
        if existing_summaries:
            response = "Voici les résumés existants pour cette conversation:\n\n"
            for filename, url in existing_summaries:
                response += f"- {filename}: {url}\n"
            
            response += "\nVoulez-vous générer un nouveau résumé? (Oui/Non)"
            
            # Set state to wait for confirmation
            self.waiting_for_confirmation = True
            self.confirmation_command = "summarize"
            self.confirmation_additional_info = additional_info
            
            return response
        else:
            # No existing summaries, generate one directly
            return self._execute_summarize(additional_info)
    
    def _execute_summarize(self, additional_info: str = None) -> str:
        """Execute the summarize operation"""
        try:
            # Import here to avoid circular imports
            from core import summarize_ppt
            
            # Generate timestamp for unique filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if acra_config.get("USE_API"):
                # Use API endpoint
                import requests
                endpoint = f"acra/{self.file_manager.chat_id}"
                if additional_info:
                    endpoint += f"?add_info={additional_info}&timestamp={timestamp}"
                else:
                    endpoint += f"?timestamp={timestamp}"
                
                url = f"{acra_config.get('API_URL')}/{endpoint}"
                response = requests.get(url)
                result = response.json() if response.status_code == 200 else {"error": "Request failed"}
            else:
                # Use direct function call
                result = summarize_ppt(self.file_manager.chat_id, additional_info, timestamp)
            
            if "error" in result:
                return f"Erreur lors de la génération du résumé: {result['error']}"
            
            # Upload result and get download URL
            upload_result = self.file_manager.upload_to_openwebui(result["summary"])
            
            if "error" in upload_result:
                return f"Résumé généré mais erreur lors du téléchargement: {upload_result['error']}"
            
            # Generate introduction if we have system prompt
            if hasattr(self, 'system_prompt') and self.system_prompt:
                introduction = model_manager.generate_introduction(self.system_prompt)
                response = f"{introduction}\n\n Le résumé de tous les fichiers a été généré avec succès.\n\n  ### URL de téléchargement: \n{upload_result.get('download_url', 'Non disponible')}"
            else:
                response = f"Le résumé de tous les fichiers a été généré avec succès.\n\n### URL de téléchargement:\n{upload_result.get('download_url', 'Non disponible')}"
            
            self.file_manager.save_file_mappings()
            return response
            
        except Exception as e:
            log.error(f"Error executing summarize: {str(e)}")
            return f"Erreur lors de la génération du résumé: {str(e)}"
    
    def handle_structure_command(self) -> str:
        """Handle /structure command"""
        try:
            if self.cached_structure is None:
                # Import here to avoid circular imports
                from core import get_slide_structure
                
                response = get_slide_structure(self.file_manager.chat_id)
                
                if "error" in response:
                    return f"Erreur lors de l'analyse de la structure: {response['error']}"
                
                # Cache the structure
                self.cached_structure = response
                # Format for display
                formatted_response = self._format_slide_data(response)
                return formatted_response
            else:
                # Use cached structure
                if isinstance(self.cached_structure, dict):
                    return self._format_slide_data(self.cached_structure)
                else:
                    return self.cached_structure
                    
        except Exception as e:
            log.error(f"Error handling structure command: {str(e)}")
            return f"Erreur lors de l'analyse de la structure: {str(e)}"
    
    def handle_generate_command(self, message: str) -> str:
        """Handle /generate command"""
        try:
            # Extract text content after the command
            text_content = message.replace("/generate", "").strip()
            if not text_content:
                return "Veuillez fournir du texte après la commande /generate pour générer un rapport."
            
            # Import here to avoid circular imports
            from core import generate_pptx_from_text
            
            # Generate timestamp for unique filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if acra_config.get("USE_API"):
                # Use API endpoint
                import requests
                endpoint = f"generate_report/{self.file_manager.chat_id}?info={text_content}&timestamp={timestamp}"
                url = f"{acra_config.get('API_URL')}/{endpoint}"
                response = requests.get(url)
                result = response.json() if response.status_code == 200 else {"error": "Request failed"}
            else:
                # Use direct function call
                result = generate_pptx_from_text(self.file_manager.chat_id, text_content, timestamp)
            
            if "error" in result:
                return f"Erreur lors de la génération du rapport: {result['error']}"
            
            # Upload result and get download URL
            upload_result = self.file_manager.upload_to_openwebui(result["summary"])
            
            if "error" in upload_result:
                return f"Rapport généré mais erreur lors du téléchargement: {upload_result['error']}"
            
            self.file_manager.save_file_mappings()
            return f"Le rapport a été généré avec succès à partir du texte fourni.\n\n### URL de téléchargement:\n{upload_result.get('download_url', 'Non disponible')}"
            
        except Exception as e:
            log.error(f"Error handling generate command: {str(e)}")
            return f"Erreur lors de la génération du rapport: {str(e)}"
    
    def handle_clear_command(self, message: str) -> str:
        """Handle /clear command"""
        try:
            # Preserve current chat ID
            preserve_ids = [self.file_manager.chat_id] if self.file_manager.chat_id else []
            
            # Extract additional IDs to preserve if specified
            if " " in message:
                additional_ids = message.split(" ", 1)[1].strip().split()
                if additional_ids:
                    preserve_ids.extend(additional_ids)
            
            # Run cleanup
            cleanup_result = cleanup_orphaned_folders()
            
            # Reset state
            self.reset_state()
            self.file_manager.file_id_mapping = {}
            
            return f"Nettoyage terminé!\n\nProtégés: {preserve_ids}\nRésultat: {cleanup_result.get('message', 'Nettoyage effectué')}"
            
        except Exception as e:
            log.error(f"Error handling clear command: {str(e)}")
            return f"Une erreur s'est produite lors du nettoyage: {str(e)}"
    
    def handle_merge_command(self) -> str:
        """Handle /merge command"""
        try:
            # Import here to avoid circular imports
            from services import merge_pptx
            
            output_merge = os.path.join(acra_config.output_folder, self.file_manager.chat_id, "merged")
            input_merge = acra_config.get_conversation_upload_folder(self.file_manager.chat_id)
            
            merge_result = merge_pptx(input_merge, output_merge)
            
            if "error" in merge_result:
                return f"Erreur lors de la fusion des fichiers: {merge_result['error']}"
            
            # Get the merged file and upload to OpenWebUI
            merged_file = merge_result.get("merged_file")
            if merged_file:
                upload_result = self.file_manager.upload_to_openwebui(merged_file)
                if "error" in upload_result:
                    return f"Les fichiers ont été fusionnés avec succès, mais une erreur s'est produite lors de la génération du lien de téléchargement: {upload_result['error']}"
                else:
                    self.file_manager.save_file_mappings()
                    return f"Les fichiers ont été fusionnés avec succès.\n\n### URL de téléchargement:\n{upload_result.get('download_url', 'Non disponible')}"
            else:
                return "Les fichiers ont été fusionnés avec succès, mais le chemin du fichier fusionné n'a pas été trouvé."
                
        except Exception as e:
            log.error(f"Error handling merge command: {str(e)}")
            return f"Erreur lors de la fusion des fichiers: {str(e)}"
    
    def handle_regroup_command(self) -> str:
        """Handle /regroup command"""
        try:
            # Get structure data
            if self.cached_structure is None:
                from core import get_slide_structure
                structure_result = get_slide_structure(self.file_manager.chat_id)
                if "error" in structure_result:
                    return f"Erreur lors de l'analyse de la structure: {structure_result['error']}"
            else:
                if isinstance(self.cached_structure, str):
                    from core import get_slide_structure
                    structure_result = get_slide_structure(self.file_manager.chat_id)
                    if "error" in structure_result:
                        return f"Erreur lors de l'analyse de la structure: {structure_result['error']}"
                else:
                    structure_result = self.cached_structure
            
            if not isinstance(structure_result, dict) or "projects" not in structure_result:
                return f"Erreur: structure de données invalide. Type: {type(structure_result)}"
            
            # Get project grouping suggestions from LLM
            project_names = list(structure_result["projects"].keys())
            grouping_response = model_manager.generate_project_grouping(project_names)
            
            try:
                groups_to_merge = extract_json(grouping_response)
                if not isinstance(groups_to_merge, list):
                    groups_to_merge = []
            except:
                log.warning("Could not extract valid JSON from LLM response")
                groups_to_merge = []
            
            log.info(f"Groups to merge: {groups_to_merge}")
            
            # Process regrouping
            new_structure = self._process_regrouping(structure_result, groups_to_merge)
            
            # Generate PowerPoint with regrouped data
            result = self._generate_regrouped_powerpoint(new_structure)
            
            # Update cached structure
            self.cached_structure = new_structure
            
            return result
            
        except Exception as e:
            log.error(f"Error handling regroup command: {str(e)}")
            return f"Erreur lors de la réorganisation des données: {str(e)}"
    
    def _format_slide_data(self, data: dict) -> str:
        """Format slide structure data for display"""
        if not data:
            return "Aucun fichier PPTX fourni."
        
        projects = data.get("projects", {})
        if not projects:
            return "Aucun projet trouvé dans les fichiers analysés."
        
        metadata = data.get("metadata", {})
        processed_files = metadata.get("processed_files", 0)
        upcoming_events = data.get("upcoming_events", {})
        
        def format_project_hierarchy(project_name, content, level=0):
            output = ""
            indent = "  " * level
            
            if level == 0:
                output += f"{indent}🔶 **{project_name}**\n"
            elif level == 1:
                output += f"{indent}📌 **{project_name}**\n"
            else:
                output += f"{indent}📎 *{project_name}*\n"
            
            if "information" in content and content["information"]:
                info_lines = content["information"].split('\n')
                for line in info_lines:
                    if line.strip():
                        output += f"{indent}- {line}\n"
                output += "\n"
            
            if "critical" in content and content["critical"]:
                output += f"{indent}- 🔴 **Alertes Critiques:**\n"
                for alert in content["critical"]:
                    output += f"{indent}  - {alert}\n"
                output += "\n"
            
            if "small" in content and content["small"]:
                output += f"{indent}- 🟡 **Alertes à surveiller:**\n"
                for alert in content["small"]:
                    output += f"{indent}  - {alert}\n"
                output += "\n"
            
            if "advancements" in content and content["advancements"]:
                output += f"{indent}- 🟢 **Avancements:**\n"
                for advancement in content["advancements"]:
                    output += f"{indent}  - {advancement}\n"
                output += "\n"
            
            for key, value in content.items():
                if isinstance(value, dict) and key not in ["information", "critical", "small", "advancements"]:
                    output += format_project_hierarchy(key, value, level + 1)
            
            return output
        
        result = f"📊 **Synthèse globale de {processed_files} fichier(s) analysé(s)**\n\n"
        
        for project_name, project_content in projects.items():
            result += format_project_hierarchy(project_name, project_content)
        
        if upcoming_events:
            result += "\n\n📅 **Événements à venir par service:**\n\n"
            for service, events in upcoming_events.items():
                if events:
                    result += f"- **{service}:**\n"
                    for event in events:
                        result += f"  - {event}\n"
                    result += "\n"
        else:
            result += "\n\n📅 **Événements à venir:** Aucun événement particulier prévu.\n"
        
        return result.strip()
    
    def _process_regrouping(self, structure_result: dict, groups_to_merge: list) -> dict:
        """Process the regrouping of projects"""
        new_structure = json.loads(json.dumps(structure_result))
        
        for group in groups_to_merge:
            if not isinstance(group, list) or len(group) < 2:
                continue
            
            # Create mapping between original and cleaned names
            original_to_cleaned = {}
            cleaned_to_original = {}
            
            for original_name in group:
                cleaned_name = original_name.replace('\n', ' ').strip()
                original_to_cleaned[original_name] = cleaned_name
                cleaned_to_original[cleaned_name] = original_name
            
            cleaned_group = list(cleaned_to_original.keys())
            main_project_cleaned = min(cleaned_group, key=len)
            main_project_original = cleaned_to_original[main_project_cleaned]
            
            other_projects_cleaned = [p for p in cleaned_group if p != main_project_cleaned]
            other_projects_original = [cleaned_to_original[p] for p in other_projects_cleaned 
                                     if cleaned_to_original[p] in new_structure["projects"]]
            
            if main_project_original not in new_structure["projects"]:
                log.warning(f"Main project '{main_project_original}' not found in structure, skipping group")
                continue
            
            # Regroup other projects under main project
            for other_project_original in other_projects_original:
                if other_project_original in new_structure["projects"]:
                    try:
                        other_data = new_structure["projects"][other_project_original]
                        other_project_cleaned = original_to_cleaned[other_project_original]
                        
                        sub_name = other_project_cleaned.replace(main_project_cleaned, "").strip("_").strip()
                        if not sub_name:
                            sub_name = other_project_cleaned
                        
                        # Handle terminal vs non-terminal projects
                        if "information" in new_structure["projects"][main_project_original]:
                            main_data = {
                                "information": new_structure["projects"][main_project_original].get("information", ""),
                                "critical": new_structure["projects"][main_project_original].get("critical", []),
                                "small": new_structure["projects"][main_project_original].get("small", []),
                                "advancements": new_structure["projects"][main_project_original].get("advancements", [])
                            }
                            
                            new_structure["projects"][main_project_original] = {
                                "Général": main_data,
                                sub_name: other_data
                            }
                        else:
                            new_structure["projects"][main_project_original][sub_name] = other_data
                        
                        del new_structure["projects"][other_project_original]
                        log.info(f"Moved {other_project_original} to {main_project_original}.{sub_name}")
                        
                    except Exception as e:
                        log.error(f"Error moving project {other_project_original}: {str(e)}")
                        continue
        
        return new_structure
    
    def _generate_regrouped_powerpoint(self, new_structure: dict) -> str:
        """Generate PowerPoint with regrouped data"""
        try:
            from src.services.update_pttx_service import update_table_with_project_data
            from pptx import Presentation
            from pptx.util import Pt
            
            # Create output directory
            output_regroup = os.path.join(acra_config.output_folder, self.file_manager.chat_id, "regrouped")
            os.makedirs(output_regroup, exist_ok=True)
            
            # Generate output filename with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_regroup, f"regrouped_{timestamp}.pptx")
            
            # Create presentation from template or blank
            if os.path.exists(acra_config.template_path):
                prs = Presentation(acra_config.template_path)
            else:
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                table_shape = slide.shapes.add_table(rows=10, cols=3, left=Pt(30), top=Pt(30), width=Pt(600), height=Pt(400))
            
            # Save temporary file
            temp_path = os.path.join(output_regroup, "temp.pptx")
            prs.save(temp_path)
            
            # Update with project data
            updated_path = update_table_with_project_data(
                temp_path,
                0,  # slide index
                0,  # table shape index
                new_structure["projects"],
                output_file,
                new_structure.get("upcoming_events", {})
            )
            
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
            
            # Upload to OpenWebUI
            upload_result = self.file_manager.upload_to_openwebui(updated_path)
            
            if "error" in upload_result:
                return f"Les informations des projets ont été regroupées avec succès, mais une erreur s'est produite lors de la génération du lien de téléchargement: {upload_result['error']}"
            
            self.file_manager.save_file_mappings()
            return f"Les informations des projets ont été regroupées avec succès.\n\n### URL de téléchargement:\n{upload_result.get('download_url', 'Non disponible')}"
            
        except Exception as e:
            log.error(f"Error generating regrouped PowerPoint: {str(e)}")
            return f"Erreur lors de la génération de la présentation PowerPoint: {str(e)}" 