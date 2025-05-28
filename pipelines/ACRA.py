import json
import os
import sys
import shutil
import requests
import sqlite3
import time
import datetime
from typing import List, Generator, Dict, Any, Optional, Tuple
from dotenv import load_dotenv

# Add src to path for imports
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

# Import core functions
from core import summarize_ppt, get_slide_structure, generate_pptx_from_text
from services import merge_pptx

# Import utilities
from OLLibrary.utils.text_service import remove_tags_no_keep
from OLLibrary.utils.log_service import setup_logging, get_logger
from OLLibrary.utils.json_service import extract_json

# Import service classes and config
from config_pipeline import acra_config
from services.file_manager import FileManager
from services.command_handler import CommandHandler
from services.model_manager import model_manager

# Set up logging
setup_logging(app_name="ACRA_Pipeline")
log = get_logger(__name__)

UPLOAD_FOLDER = acra_config.upload_folder
OUTPUT_FOLDER = acra_config.output_folder
MAPPINGS_FOLDER = acra_config.mappings_folder

class Pipeline:
    def __init__(self):
        log.info("Initializing ACRA Pipeline")
        
        # Initialize configuration and ensure directories exist
        acra_config.ensure_directories()
        
        # State tracking - chat_id needs to be initialized before FileManager
        self.chat_id = "" 
        self.current_chat_id = ""
        
        # Initialize file manager
        self.file_manager = FileManager(chat_id=self.chat_id)
        
        # Initialize command handler, passing the file_manager instance
        self.command_handler = CommandHandler(file_manager=self.file_manager)
        
        self.system_prompt = ""
        self.message_id = 0
        self.file_path_list = []
        self.last_response = None
        
        self.cached_structure = None

        # Validate required configuration (OPENWEBUI_API_KEY is checked by acra_config or FileManager if needed)
        if not acra_config.get("OPENWEBUI_API_KEY"):
            log.error("OPENWEBUI_API_KEY is not set in acra_config")
            raise ValueError("OPENWEBUI_API_KEY is not set")
        
        log.info("ACRA Pipeline initialized successfully")

    def generate_report(self, foldername, info):
        """
        Generate a report from provided text using API or direct function call.
        Creates a new file with unique timestamp for each call.
        
        Args:
            foldername (str): Folder name to store the report. Should align with chat_id.
            info (str): Text to analyze for report generation
            
        Returns:
            dict: Request result with download URL
        """
        log.info(f"Generating report for chat: {self.file_manager.chat_id}") # Use file_manager.chat_id
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log.info(f"Creating report with timestamp: {timestamp}")
        
        if acra_config.get("USE_API"): 
            log.info("Using API endpoint to generate report")
            # Ensure foldername here is self.file_manager.chat_id for consistency
            endpoint = f"generate_report/{self.file_manager.chat_id}?info={info}&timestamp={timestamp}"
            result = self._fetch_api(endpoint)
            if "error" in result:
                return result
            return self.file_manager.upload_to_openwebui(result["summary"])

        log.info("Using direct function call to generate report")
        # Ensure foldername passed to generate_pptx_from_text is self.file_manager.chat_id
        result = generate_pptx_from_text(self.file_manager.chat_id, info, timestamp)
        if "error" in result:
            return result
            
        upload_result = self.file_manager.upload_to_openwebui(result["summary"])
        self.file_manager.save_file_mappings()
        return upload_result

    def reset_conversation_state(self):
        """Reset conversation-specific states"""
        log.info(f"Resetting conversation state for chat_id: {self.chat_id}")
        self.command_handler.reset_state()
        self.system_prompt = ""
        self.file_path_list = []
        self.message_id = 0


    def save_file_mappings(self):
        """Save file mappings (delegated to file manager)"""
        self.file_manager.save_file_mappings()

    def _fetch_api(self, endpoint):
        """Perform synchronous GET request"""
        url = f"{acra_config.get('API_URL')}/{endpoint}"
        log.debug(f"Fetching from: {url}")
        response = requests.get(url)
        if response.status_code != 200:
            log.error(f"API request failed: {response.status_code} - {response.text}")
        return response.json() if response.status_code == 200 else {"error": "Request failed"}

    def _post_api(self, endpoint, data=None, files=None, headers=None):
        """Perform synchronous POST request"""
        if endpoint.startswith("http"):
            url = endpoint
        else:
            url = f"{acra_config.get('API_URL')}/{endpoint}"
        
        log.debug(f"Posting to: {url}")
        response = requests.post(url, data=data, files=files, headers=headers)
        if response.status_code != 200:
            log.error(f"API POST request failed: {response.status_code} - {response.text}")
        return response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}

    def download_file_openwebui(self, file: str):
        """Upload file to OpenWebUI and get download URL (delegated to file manager)"""
        return self.file_manager.upload_to_openwebui(file)
    
    def summarize_folder(self, foldername=None, add_info=None):
        """
        Send request to summarize all PowerPoint files in a folder.
        Generates a new file with unique timestamp for each call.
        Delegates to CommandHandler._execute_summarize after confirmation logic (if any).
        This method in Pipeline might just become a simple wrapper or be removed if CommandHandler handles it.
        For now, let's assume CommandHandler's handle_summarize_command will call its _execute_summarize.
        
        Args:
            foldername (str, optional): Folder name to summarize. If None, uses chat_id.
            add_info (str, optional): Additional information to add to summary.
        Returns:
            dict: Summary operation results.
        """
        current_chat_id = foldername if foldername else self.file_manager.chat_id
        log.info(f"Request to summarize folder for chat: {current_chat_id}")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log.info(f"Creating summary with timestamp: {timestamp}")
        
        if acra_config.get("USE_API"):
            log.info("Using API endpoint to summarize folder")
            endpoint = f"acra/{current_chat_id}"
            if add_info:
                endpoint += f"?add_info={add_info}&timestamp={timestamp}"
            else:
                endpoint += f"?timestamp={timestamp}"
            result = self._fetch_api(endpoint)
            if "error" in result:
                return result
            
            self._track_source_files(current_chat_id)
            upload_result = self.file_manager.upload_to_openwebui(result["summary"])
            self.file_manager.save_file_mappings()
            return upload_result
        
        log.info("Using direct function call to summarize folder")
        result = summarize_ppt(current_chat_id, add_info, timestamp) # summarize_ppt needs chat_id
        if "error" in result:
            return result
        
        upload_result = self.file_manager.upload_to_openwebui(result["summary"])
        self._track_source_files(current_chat_id)
        self.file_manager.save_file_mappings()
        return upload_result
    
    def _track_source_files(self, foldername): # foldername should be chat_id
        """Track source files from the upload folder for the given chat_id"""
        source_folder = acra_config.get_conversation_upload_folder(foldername) # foldername is chat_id
        if os.path.exists(source_folder):
            for filename in os.listdir(source_folder):
                if filename.lower().endswith(".pptx"): # or use acra_config.get("ALLOWED_EXTENSIONS")
                    source_file_path = os.path.join(source_folder, filename)
                    abs_file_path = os.path.abspath(source_file_path)
                    # Check against file_manager's mapping
                    if abs_file_path not in self.file_manager.file_id_mapping:
                        log.info(f"Source file not yet tracked by FileManager: {abs_file_path}")
                        # Optionally, trigger an upload/mapping via file_manager here if needed
                        # self.file_manager.upload_to_openwebui(abs_file_path)

    def extract_service_name(self, filename):
        """
        Extract service name from PowerPoint filename using the model_manager.
        """
        # Delegate to model_manager
        return model_manager.extract_service_name(filename)

    def analyze_slide_structure(self, foldername=None): # foldername should be chat_id
        """
        Analyze slide structure in a folder.
        """
        # foldername logic should align with self.file_manager.chat_id
        current_chat_id = foldername if foldername else self.file_manager.chat_id
        
        if not current_chat_id: # Check if chat_id is actually set
            log.error("Chat ID is not set. Cannot analyze slide structure.")
            # Consider raising an exception or returning an error dict
            raise ValueError("Chat ID is not set. Cannot analyze slide structure.")
        
        log.info(f"Analyzing slide structure for chat: {current_chat_id}")
        
        if acra_config.get("USE_API"):
            log.info("Using API endpoint to get slide structure")
            return self._fetch_api(f"get_slide_structure/{current_chat_id}")
        
        log.info("Using direct function call to get slide structure")
        return get_slide_structure(current_chat_id) # get_slide_structure needs chat_id
    
    def format_all_slide_data(self, data: dict) -> str:
        """
        Formats slide data. This is also present in CommandHandler.
        Consider consolidating or ensuring CommandHandler uses this if it's preferred here.
        For now, keeping implementation here as it was, but flag for review.
        """
        # Si data est None ou vide, renvoyer un message d'erreur
        if not data:
            return "Aucun fichier PPTX fourni."
            
        # Utiliser directement les données fournies sans modifier le cache
        structure_to_process = data
        
        # Vérifier si nous avons des projets
        projects = structure_to_process.get("projects", {})
        if not projects:
            return "Aucun projet trouvé dans les fichiers analysés."
            
        # Récupérer les métadonnées et événements à venir
        metadata = structure_to_process.get("metadata", {})
        processed_files = metadata.get("processed_files", 0)
        upcoming_events = structure_to_process.get("upcoming_events", {})
            
        # Fonction récursive pour afficher les projets à tous les niveaux de hiérarchie
        def format_project_hierarchy(project_name, content, level=0):
            output = ""
            indent = "  " * level
            
            # Format le nom du projet selon son niveau
            if level == 0:
                output += f"{indent}🔶 **{project_name}**\n"
            elif level == 1:
                output += f"{indent}📌 **{project_name}**\n"
            else:
                output += f"{indent}📎 *{project_name}*\n"
            
            # Ajouter les informations si elles existent
            if "information" in content and content["information"]:
                info_lines = content["information"].split('\n')
                for line in info_lines:
                    if line.strip():
                        output += f"{indent}- {line}\n"
                output += "\n"
            
            # Ajouter les alertes critiques
            if "critical" in content and content["critical"]:
                output += f"{indent}- 🔴 **Alertes Critiques:**\n"
                for alert in content["critical"]:
                    output += f"{indent}  - {alert}\n"
                output += "\n"
            
            # Ajouter les alertes à surveiller
            if "small" in content and content["small"]:
                output += f"{indent}- 🟡 **Alertes à surveiller:**\n"
                for alert in content["small"]:
                    output += f"{indent}  - {alert}\n"
                output += "\n"
            
            # Ajouter les avancements
            if "advancements" in content and content["advancements"]:
                output += f"{indent}- 🟢 **Avancements:**\n"
                for advancement in content["advancements"]:
                    output += f"{indent}  - {advancement}\n"
                output += "\n"
            
            # Traiter les sous-projets ou sous-sous-projets de façon récursive
            for key, value in content.items():
                if isinstance(value, dict) and key not in ["information", "critical", "small", "advancements"]:
                    output += format_project_hierarchy(key, value, level + 1)
            
            return output

        # Créer le résultat final
        result = ""
        
        # Afficher le nombre de présentations analysées
        result += f"📊 **Synthèse globale de {processed_files} fichier(s) analysé(s)**\n\n"
        
        # Formater chaque projet principal
        for project_name, project_content in projects.items():
            result += format_project_hierarchy(project_name, project_content)
        
        # Ajouter la section des événements à venir par service
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

    def delete_all_files(self, foldername=None): # foldername should be chat_id
        """
        Deletes all files for a given chat_id.
        Manages local file deletion and OpenWebUI file deletion.
        Uses FileManager for some operations.
        """
        current_chat_id = foldername if foldername else self.file_manager.chat_id
        log.info(f"Deleting all files for chat: {current_chat_id}")

        pptx_folder_path = acra_config.get_conversation_upload_folder(current_chat_id)
        output_folder_path = acra_config.get_conversation_output_folder(current_chat_id)
        
        deleted_local_files = 0
        
        # 1. Delete OpenWebUI files associated with this chat
        webui_result = self.delete_openwebui_files_for_chat(current_chat_id) 
        deleted_webui_files = webui_result.get("deleted_count", 0)
        log.info(f"OpenWebUI deletion result: {webui_result}")
        
        # 2. Delete local files from pptx_folder (upload folder)
        if os.path.exists(pptx_folder_path):
            for item in os.listdir(pptx_folder_path):
                item_path = os.path.join(pptx_folder_path, item)
                if os.path.isfile(item_path):
                    try:
                        os.remove(item_path)
                        deleted_local_files += 1
                        log.info(f"Deleted local file: {item_path}")
                    except Exception as e:
                        log.error(f"Error deleting local file {item_path}: {e}")
            # Optionally remove the folder itself if empty, or let cleanup handle it
            # shutil.rmtree(pptx_folder_path, ignore_errors=True) 
        
        # 3. Delete local files from OUTPUT folder
        if os.path.exists(output_folder_path):
            for item in os.listdir(output_folder_path):
                item_path = os.path.join(output_folder_path, item)
                if os.path.isfile(item_path) and item.endswith(".pptx"): # Only PPTX, leave mappings
                    try:
                        os.remove(item_path)
                        deleted_local_files += 1
                        log.info(f"Deleted local output file: {item_path}")
                    except Exception as e:
                        log.error(f"Error deleting local output file {item_path}: {e}")
            # Optionally remove the folder itself
            # shutil.rmtree(output_folder_path, ignore_errors=True)

        # 4. Clear and delete mapping file for this conversation via FileManager
        self.file_manager.file_id_mapping.clear() # Clear current instance's map
        mapping_file_path = acra_config.get_mapping_file_path(current_chat_id)
        if os.path.exists(mapping_file_path):
            try:
                os.remove(mapping_file_path)
                log.info(f"Deleted mapping file: {mapping_file_path}")
            except Exception as e:
                    log.error(f"Error deleting mapping file {mapping_file_path}: {e}")
        
        # Reset pipeline state associated with files
        self.file_path_list = []
        self.command_handler.cached_structure = None # Or self.cached_structure if still used by Pipeline

        message = f"{deleted_local_files} fichiers locaux supprimés. {deleted_webui_files} fichiers supprimés d'OpenWebUI."
        return {"message": message, "deleted_local_files": deleted_local_files, "deleted_webui_files": deleted_webui_files}

    def get_files_in_folder(self, foldername=None): # foldername should be chat_id
        """
        Retrieves list of PPTX files in the upload folder for the given chat_id.
        """
        current_chat_id = foldername if foldername else self.file_manager.chat_id
        return self.file_manager.get_conversation_files(file_extensions=[".pptx"])

    def get_active_conversation_ids(self):
        """
        Retrieves active conversation IDs from OpenWebUI database.
        This method relies on direct DB access and specific table/column names.
        It's a critical part of cleanup.
        Uses acra_config for DB path.
        """
        conversation_ids = []
        db_path = acra_config.get("OPENWEBUI_DB_PATH")
        log.info(f"Attempting to access OpenWebUI database at: {db_path}")

        try:
            if not os.path.exists(db_path):
                log.error(f"OpenWebUI database not found at {db_path}")
                alt_paths = ["./webui.db", "/app/webui.db", "/app/open-webui/webui.db"]
                for path_attempt in alt_paths:
                    abs_path_attempt = os.path.abspath(path_attempt)
                    log.info(f"Attempting alternative DB path: {abs_path_attempt}")
                    if os.path.exists(abs_path_attempt):
                        log.info(f"Found database at alternative path: {abs_path_attempt}")
                        db_path = abs_path_attempt
                        break
                else:
                    log.error("Could not find OpenWebUI database in any expected location")
                    return self.get_all_existing_chat_folders() # Fallback
            
            log.info(f"Connecting to SQLite database at: {db_path}")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='chat'")
            if not cursor.fetchone():
                log.error("Table 'chat' not found in database")
                conn.close()
                return self.get_all_existing_chat_folders()
            
            cursor.execute("PRAGMA table_info(chat)")
            columns = [col[1] for col in cursor.fetchall()]
            
            if "deleted_at" in columns:
                log.info("Using deleted_at column to filter active chats")
                cursor.execute("SELECT id FROM chat WHERE deleted_at IS NULL OR deleted_at = ''")
            else:
                log.info("deleted_at column not found, getting all chat IDs")
                cursor.execute("SELECT id FROM chat")
                
            rows = cursor.fetchall()
            base_conversation_ids = [str(row[0]) for row in rows] # Ensure string IDs
            log.info(f"Found {len(base_conversation_ids)} conversations in database: {base_conversation_ids}")
            
            # Check messages table (original logic kept)
            try:
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='message'") # Corrected table name
                if cursor.fetchone(): # Check if 'message' table exists
                    log.info("Checking messages table for active chats")
                    # Query for chat_ids that have messages
                    cursor.execute("SELECT DISTINCT chat_id FROM message WHERE chat_id IS NOT NULL")
                    message_chat_ids_rows = cursor.fetchall()
                    message_chat_ids = [str(row[0]) for row in message_chat_ids_rows if row[0]] # Ensure string and not None
                    
                    log.info(f"Found {len(message_chat_ids)} chats with messages: {message_chat_ids}")
                    
                    for chat_id_from_msg in message_chat_ids:
                        if chat_id_from_msg not in base_conversation_ids:
                            base_conversation_ids.append(chat_id_from_msg)
                            log.info(f"Added chat with messages that was not in main list: {chat_id_from_msg}")
            except Exception as e:
                log.error(f"Error checking messages table: {str(e)}")
            
            conn.close()
            conversation_ids = list(set(base_conversation_ids)) # Ensure unique
            log.info(f"Final active conversations: {conversation_ids}")

        except Exception as e:
            log.error(f"Error retrieving conversation IDs from database: {str(e)}")
            log.exception("Database access exception details:")
            return self.get_all_existing_chat_folders()
        
        if not conversation_ids:
            log.warning("No active conversations found in database! Using existing folders as fallback.")
            conversation_ids = self.get_all_existing_chat_folders()
        
        current_pipeline_chat_id = self.file_manager.chat_id # Use FileManager's chat_id
        if current_pipeline_chat_id and current_pipeline_chat_id not in conversation_ids:
            log.warning(f"Current chat_id {current_pipeline_chat_id} not found in active list! Adding it.")
            conversation_ids.append(current_pipeline_chat_id)
        
        return list(set(conversation_ids)) # Ensure unique again after adding current
    
    def get_all_existing_chat_folders(self):
        """
        Fallback: Gets IDs of all existing chat folders from UPLOAD, OUTPUT, MAPPINGS.
        Uses acra_config for paths.
        """
        folder_ids = set()
        upload_dir = acra_config.upload_folder
        output_dir = acra_config.output_folder
        mappings_dir = acra_config.mappings_folder
        
        for base_dir in [upload_dir, output_dir]:
            if os.path.exists(base_dir):
                try:
                    for folder_name in os.listdir(base_dir):
                        if os.path.isdir(os.path.join(base_dir, folder_name)):
                            folder_ids.add(folder_name)
                except Exception as e:
                    log.error(f"Error listing {base_dir}: {str(e)}")
        
        if os.path.exists(mappings_dir):
            try:
                for filename in os.listdir(mappings_dir):
                    if filename.endswith("_file_mappings.json"):
                        chat_id = filename.split("_file_mappings.json")[0]
                        folder_ids.add(chat_id)
            except Exception as e:
                log.error(f"Error listing {mappings_dir}: {str(e)}")
        
        folder_list = list(folder_ids)
        log.info(f"Found {len(folder_list)} existing chat folders as safety fallback: {folder_list}")
        return folder_list

    def cleanup_orphaned_conversations(self):
        """
        Cleans orphaned conversation folders and files.
        Relies on get_active_conversation_ids.
        Uses acra_config for paths.
        """
        log.info("Starting cleanup of orphaned conversations")
        active_conversations = self.get_active_conversation_ids()
        log.info(f"Active conversations from DB/fallback: {active_conversations}")

        if not active_conversations: # Should be handled by get_active_conversation_ids returning fallback
            log.warning("No active conversations determined. Skipping cleanup for safety.")
            return {"status": "warning", "message": "No active conversations - No cleanup performed for safety", "action": "none"}

        # Safety: If very few active conversations, could be DB error.
        if len(active_conversations) < 2 and self.file_manager.chat_id not in active_conversations:
             # If the only "active" is not even the current one, it's risky.
            log.warning(f"Only {len(active_conversations)} active_conversations. Current chat '{self.file_manager.chat_id}' is not in it. Skipping cleanup.")
            return {
                "status": "warning", 
                "message": f"Only {len(active_conversations)} active conversation(s) found, current chat not among them. Skipping cleanup for safety.", 
                "action": "none"
            }
        
        # Ensure current chat_id (from file_manager) is always preserved
        current_fm_chat_id = self.file_manager.chat_id
        if current_fm_chat_id and current_fm_chat_id not in active_conversations:
            log.info(f"Adding current FileManager chat_id to active list for preservation: {current_fm_chat_id}")
            active_conversations.append(current_fm_chat_id)
            active_conversations = list(set(active_conversations)) # Ensure unique
        
        deleted_folders_count = 0
        deleted_files_count = 0 # Local files in those folders
        deleted_mapping_files = 0
        
        upload_dir = acra_config.upload_folder
        output_dir = acra_config.output_folder
        mappings_dir = acra_config.mappings_folder

        # Iterate over folders in UPLOAD_FOLDER, OUTPUT_FOLDER, MAPPINGS_FOLDER
        for base_folder_path, type_of_folder in [
            (upload_dir, "upload"), 
            (output_dir, "output"),
            (mappings_dir, "mappings_file") # Special handling for mapping files
        ]:
            if not os.path.exists(base_folder_path):
                log.info(f"{type_of_folder} directory not found at {base_folder_path}, skipping.")
                continue

            log.info(f"Checking for orphaned items in {base_folder_path} ({type_of_folder})")
            for item_name in os.listdir(base_folder_path):
                item_path = os.path.join(base_folder_path, item_name)
                chat_id_candidate = None

                if type_of_folder == "mappings_file":
                    if item_name.endswith("_file_mappings.json") and os.path.isfile(item_path):
                        chat_id_candidate = item_name.split("_file_mappings.json")[0]
                    else:
                        continue # Not a mapping file we manage
                elif os.path.isdir(item_path): # For upload and output folders
                    chat_id_candidate = item_name
                else: # Not a directory in upload/output
                    continue 
                
                if chat_id_candidate and chat_id_candidate not in active_conversations:
                    log.info(f"Orphaned item found: {item_path} (chat_id: {chat_id_candidate})")
                    try:
                        if type_of_folder == "mappings_file":
                            os.remove(item_path)
                            deleted_mapping_files += 1
                            log.info(f"Deleted orphaned mapping file: {item_path}")
                        else: # It's an upload or output directory
                            # Count files before deleting directory
                            num_files_in_dir = 0
                            for root, _, files_in_subdir in os.walk(item_path):
                                num_files_in_dir += len(files_in_subdir)
                            
                            shutil.rmtree(item_path, ignore_errors=True)
                            deleted_folders_count += 1
                            deleted_files_count += num_files_in_dir
                            log.info(f"Deleted orphaned folder: {item_path} (contained {num_files_in_dir} files)")
                    except Exception as e:
                        log.error(f"Failed to delete orphaned item {item_path}: {e}")
        
        result = {
            "status": "success",
            "deleted_folders": deleted_folders_count,
            "deleted_files_in_folders": deleted_files_count,
            "deleted_mapping_files": deleted_mapping_files,
            "active_conversations_checked": len(active_conversations),
            "action": "cleanup_orphaned_conversations"
        }
        log.info(f"Orphaned conversation cleanup results: {result}")
        return result

    async def inlet(self, body: dict, user: dict) -> dict:
        log.info(f"Received body: {body}")
        metadata = body.get("metadata", {})
        log.info(f"Metadata: {metadata}")
        
        current_fm_chat_id = self.file_manager.chat_id
        log.info(f"Current state - Pipeline.chat_id: '{self.chat_id}', FileManager.chat_id: '{current_fm_chat_id}'")
        
        new_chat_id_from_request = metadata.get("chat_id")
        
        if new_chat_id_from_request:
            log.info(f"Chat_id from request metadata: '{new_chat_id_from_request}'")
            
            if new_chat_id_from_request != current_fm_chat_id:
                log.info(f"*** CHAT ID CHANGING *** from FileManager current '{current_fm_chat_id}' to new '{new_chat_id_from_request}'")
                
                # FileManager's set_chat_id handles saving old mappings and loading new ones.
                self.file_manager.set_chat_id(new_chat_id_from_request)
                
                # Update Pipeline's chat_id as well, though FileManager is the primary owner now.
                self.chat_id = new_chat_id_from_request 
                
                # Reset pipeline states specific to a conversation
                self.reset_conversation_state() # This now delegates to command_handler.reset_state()
                
                # Trigger cleanup of orphaned conversations (now that active chats might have changed)
                # Consider if this should be done before or after setting new chat_id.
                # Doing it after ensures the new chat_id is preserved.
                self._cleanup_old_chat(old_chat_id=current_fm_chat_id, new_chat_id=new_chat_id_from_request)

            elif not current_fm_chat_id: # First time FileManager gets a chat_id
                log.info(f"Setting initial FileManager chat_id to {new_chat_id_from_request}")
                self.file_manager.set_chat_id(new_chat_id_from_request)
                self.chat_id = new_chat_id_from_request # Keep Pipeline's chat_id in sync
        else:
            log.warning("No chat_id found in request metadata!")

        files_metadata = metadata.get("files", [])
        if files_metadata:
            self.command_handler.cached_structure = None # Reset cache in CommandHandler
            
            for file_entry in files_metadata:
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                openwebui_file_id = file_data.get("id", "N/A") # This is OpenWebUI's internal file ID

                openwebui_uploads_dir = acra_config.get("OPENWEBUI_UPLOADS", "open-webui/uploads") # From acra_config
                source_path_in_openwebui_volume = os.path.join(openwebui_uploads_dir, f"{openwebui_file_id}_{filename}")
                
                if not os.path.exists(source_path_in_openwebui_volume):
                    log.error(f"Source file from OpenWebUI not found at: {source_path_in_openwebui_volume}")
                    # Try an alternative common path if the primary one fails
                    alt_source_path = os.path.join("uploads", f"{openwebui_file_id}_{filename}") # Relative to CWD
                    if os.path.exists(alt_source_path):
                        source_path_in_openwebui_volume = alt_source_path
                        log.info(f"Found source file at alternative path: {alt_source_path}")
                    else:
                        log.error(f"Also not found at alternative: {alt_source_path}. Skipping file.")
                        continue
                copied_file_path = self.file_manager.copy_uploaded_file(
                    source_path=source_path_in_openwebui_volume, 
                    filename=f"{openwebui_file_id}_{filename}"
                )
                
                abs_copied_path = os.path.abspath(copied_file_path)
                if abs_copied_path not in self.file_manager.file_id_mapping:
                    self.file_manager.file_id_mapping[abs_copied_path] = openwebui_file_id
                    log.info(f"Manually added mapping for copied file: {abs_copied_path} -> {openwebui_file_id}")
                
                service_name = model_manager.extract_service_name(filename) # Use model_manager
                log.info(f"File {filename} (OpenWebUI ID: {openwebui_file_id}) identified as service: {service_name}")
                
            # Analyze structure using current chat_id from file_manager
            structure_response_data = self.analyze_slide_structure(self.file_manager.chat_id) 
            
            if isinstance(structure_response_data, dict) and "error" in structure_response_data:
                display_response = f"Erreur lors de l'analyse de la structure: {structure_response_data['error']}"
                self.command_handler.cached_structure = {"error": display_response} # Cache error state
            elif isinstance(structure_response_data, dict):
                self.command_handler.cached_structure = structure_response_data # Cache in CommandHandler
                display_response = self.format_all_slide_data(structure_response_data) # Format for system prompt
            else: # Should not happen if analyze_slide_structure is consistent
                display_response = "Erreur: Type de réponse inattendu de l'analyse de structure."
                self.command_handler.cached_structure = {"error": display_response}

            self.system_prompt = f"# Voici les informations des fichiers PPTX toutes les informations sont importantes pour la compréhension du message de l\'utilisateur et les données sont triées : \n\n{display_response}\n\n# voici le message de l\'utilisateur : "
            
            self.file_manager.save_file_mappings() # Save any new manual mappings
        
        return body
    
    def _cleanup_old_chat(self, old_chat_id: Optional[str], new_chat_id: str):
        """
        Handle cleanup of old chat folders and files when chat ID changes.
        This version is simplified, relying on cleanup_orphaned_conversations for robustness.
        """
        if not old_chat_id or old_chat_id == new_chat_id:
            log.info(f"No old chat to clean up, or old is same as new ({old_chat_id} -> {new_chat_id}).")
            # Still run general orphan cleanup, as other chats might be orphaned.
            cleanup_result = self.cleanup_orphaned_conversations()
            log.info(f"General orphaned conversation cleanup result: {cleanup_result}")
            return
            
        log.info(f"Specific cleanup for old chat ID: {old_chat_id} (new is {new_chat_id})")
        
        # Get active chats. The new_chat_id should be among them.
        # cleanup_orphaned_conversations will use this list to decide what's truly orphaned.
        active_chats = self.get_active_conversation_ids() 
        log.info(f"Active chats for cleanup context: {active_chats}")

        if old_chat_id not in active_chats:
            log.info(f"Old chat ID {old_chat_id} is not in the active list. It's a candidate for deletion by general cleanup.")
            # Delete OpenWebUI files specifically for this old_chat_id if it's truly gone.
            # The delete_openwebui_files_for_chat checks if files are used by *other* active chats.
            webui_del_result = self.delete_openwebui_files_for_chat(old_chat_id)
            log.info(f"OpenWebUI file deletion result for specifically orphaned chat {old_chat_id}: {webui_del_result}")
            
            # Local folders (upload, output, mappings) for old_chat_id will be handled by
            # cleanup_orphaned_conversations if old_chat_id is not in active_conversations.
        else:
            log.info(f"Old chat ID {old_chat_id} is still considered active. No specific targeted deletion.")

        # Run the general cleanup process which handles all non-active chats
        cleanup_result = self.cleanup_orphaned_conversations()
        log.info(f"General orphaned conversation cleanup result during chat switch: {cleanup_result}")

    def get_existing_summaries(self, folder_name=None): # folder_name is chat_id
        """
        Get list of existing summary files for the current chat_id (via file_manager).
        Uploads them to OpenWebUI via file_manager to get URLs.
        """
        current_chat_id = folder_name if folder_name else self.file_manager.chat_id
        if not current_chat_id:
            log.warning("No chat_id available in get_existing_summaries.")
            return []
            
        log.info(f"Getting existing summaries for chat: {current_chat_id}")
        summaries = [] # List of (display_name, url)

        output_dir_for_chat = acra_config.get_conversation_output_folder(current_chat_id)
        upload_dir_for_chat = acra_config.get_conversation_upload_folder(current_chat_id)

        # Check OUTPUT folder
        if os.path.exists(output_dir_for_chat):
            for filename in os.listdir(output_dir_for_chat):
                if filename.lower().endswith(".pptx"): # Or use acra_config allowed extensions
                    file_path = os.path.join(output_dir_for_chat, filename)
                    upload_result = self.file_manager.upload_to_openwebui(file_path)
                    if "download_url" in upload_result:
                        summaries.append((f"OUTPUT/{filename}", upload_result["download_url"]))
            else:
                        log.error(f"Failed to get OpenWebUI URL for summary {file_path}: {upload_result.get('error')}")
        
        # Check UPLOAD folder (for reports/summaries that might be saved there)
        if os.path.exists(upload_dir_for_chat):
            for filename in os.listdir(upload_dir_for_chat):
                # Heuristic to identify summary-like files in upload folder
                if filename.lower().endswith(".pptx") and ("_summary_" in filename.lower() or "_report_" in filename.lower() or "regrouped_" in filename.lower()):
                    file_path = os.path.join(upload_dir_for_chat, filename)
                    upload_result = self.file_manager.upload_to_openwebui(file_path)
                    if "download_url" in upload_result:
                        summaries.append((f"Uploads/{filename}", upload_result["download_url"])) # Indicate source
            else:
                        log.error(f"Failed to get OpenWebUI URL for file {file_path} in uploads: {upload_result.get('error')}")
            
        if summaries: # If any files were (re-)uploaded, their mappings might have been updated
            self.file_manager.save_file_mappings()
                
        log.info(f"Found {len(summaries)} existing summaries/reports for chat {current_chat_id}: {summaries}")
        return summaries

    def delete_file(self, file_path, update_mapping=True):
        """
        Deletes a specific file. Delegates to FileManager.
        The file_path here should be an absolute path to a local file.
        """
        log.info(f"Pipeline.delete_file called for: {file_path}")
        return self.file_manager.delete_file(file_path=file_path, update_mapping=update_mapping)

    def cleanup_orphaned_mappings(self):
        """
        Cleans mappings in FileManager that point to non-existent files.
        Delegates to FileManager.
        """
        log.info("Pipeline.cleanup_orphaned_mappings called.")
        return self.file_manager.cleanup_orphaned_mappings()

    def force_cleanup_old_folders(self, exclude_chat_ids=None):
        """
        Simplified method to clean folders not matching excluded chat IDs.
        This is a more aggressive cleanup.
        It should use get_active_conversation_ids for safety.
        """
        if exclude_chat_ids is None:
            exclude_chat_ids = []
            
        # Always preserve the current chat_id from FileManager
        current_fm_chat_id = self.file_manager.chat_id
        if current_fm_chat_id and current_fm_chat_id not in exclude_chat_ids:
            exclude_chat_ids.append(current_fm_chat_id)
        
        log.info(f"Force cleaning folders, initial exclusions: {exclude_chat_ids}")
        
        # Get all truly active conversations from DB/fallback as the source of truth for preservation
        safe_to_preserve_ids = self.get_active_conversation_ids()
        log.info(f"Active chats from DB/fallback (for safety): {safe_to_preserve_ids}")
        
        # Combine exclude_chat_ids with safe_to_preserve_ids
        final_exclude_ids = list(set(exclude_chat_ids + safe_to_preserve_ids))
        log.info(f"Final exclusion list for force_cleanup: {final_exclude_ids}")
        
        deleted_folders_count = 0
        deleted_files_in_folders_count = 0
        deleted_webui_files_count = 0
        
        folders_actually_cleaned = []

        # Directories to check: UPLOAD_FOLDER, OUTPUT_FOLDER, MAPPINGS_FOLDER
        dirs_to_scan = {
            acra_config.upload_folder: "upload_dir",
            acra_config.output_folder: "output_dir",
            acra_config.mappings_folder: "mappings_dir_files" # Special handling for files here
        }

        for dir_path, dir_type in dirs_to_scan.items():
            if not os.path.exists(dir_path):
                log.warning(f"Directory {dir_path} for {dir_type} does not exist. Skipping.")
                continue

            for item_name in os.listdir(dir_path):
                item_path = os.path.join(dir_path, item_name)
                chat_id_of_item = None

                if dir_type == "mappings_dir_files":
                    if item_name.endswith("_file_mappings.json") and os.path.isfile(item_path):
                        chat_id_of_item = item_name.split("_file_mappings.json")[0]
                    else:
                        continue # Not a mapping file
                elif os.path.isdir(item_path): # For upload and output dirs
                    chat_id_of_item = item_name
                else: # Not a dir in upload/output
                    continue
            
                if chat_id_of_item and chat_id_of_item not in final_exclude_ids:
                    log.info(f"Force cleanup target: {item_path} (chat_id: {chat_id_of_item})")
                    
                    # Safety: Check for recent modification (less than CLEANUP_RETENTION_HOURS from config)
                    try:
                        retention_hours = acra_config.get("CLEANUP_RETENTION_HOURS", 24)
                        item_mod_time = os.path.getmtime(item_path)
                        if (time.time() - item_mod_time) / 3600 < retention_hours:
                            log.warning(f"Item {item_path} was modified recently. Skipping force cleanup for this item.")
                            # Add to final_exclude_ids to prevent re-processing in this run
                            final_exclude_ids.append(chat_id_of_item)
                            final_exclude_ids = list(set(final_exclude_ids))
                            continue
                    except Exception as e:
                        log.error(f"Error checking modification time for {item_path}: {e}. Skipping for safety.")
                        continue # Skip if can't check mod time

                    folders_actually_cleaned.append(chat_id_of_item)
                    
                    # 1. Delete OpenWebUI files for this chat_id_of_item
                    webui_del_res = self.delete_openwebui_files_for_chat(chat_id_of_item)
                    deleted_webui_files_count += webui_del_res.get("deleted_count", 0)
                    
                    # 2. Delete local item
                    try:
                        if dir_type == "mappings_dir_files":
                            os.remove(item_path)
                            log.info(f"Force deleted mapping file: {item_path}")
                        else: # Upload or output directory
                            files_inside = 0
                            for _, _, files_in_subdir in os.walk(item_path): files_inside += len(files_in_subdir)
                            shutil.rmtree(item_path, ignore_errors=True)
                            deleted_folders_count += 1
                            deleted_files_in_folders_count += files_inside
                            log.info(f"Force deleted folder: {item_path} (contained {files_inside} files)")
                    except Exception as e:
                        log.error(f"Error during force deletion of {item_path}: {e}")
            
        result = {
            "status": "success",
            "deleted_folders": deleted_folders_count,
            "deleted_files_in_folders": deleted_files_in_folders_count,
            "deleted_webui_files": deleted_webui_files_count,
            "preserved_chats_total": final_exclude_ids, # Show all that were ultimately preserved
            "cleaned_chat_ids": list(set(folders_actually_cleaned))
        }
        log.info(f"Force cleanup results: {result}")
        return result

    def delete_file_from_openwebui(self, file_id, active_files_mapping=None):
        """
        Deletes a file in OpenWebUI via API.
        This is a helper and needs OPENWEBUI_API_URL and OPENWEBUI_API_KEY from acra_config.
        The active_files_mapping is crucial for safety.
        """
        try:
            if not file_id:
                log.warning("Attempted to delete OpenWebUI file without an ID.")
                return False
                
            # Safety check: Is this file_id used by any *other* active conversations?
            # active_files_mapping should be {file_id: [list_of_active_chat_ids_using_it]}
            if active_files_mapping and file_id in active_files_mapping:
                log.info(f"File {file_id} is in active_files_mapping, indicating potential use by other active chats. Deferring to caller logic.")
                pass


            url = f"{acra_config.get('OPENWEBUI_API_URL')}files/{file_id}" # Use acra_config
            headers = {
                "accept": "application/json",
                "Authorization": f"Bearer {acra_config.get('OPENWEBUI_API_KEY')}" # Use acra_config
            }
            
            log.info(f"Attempting to delete OpenWebUI file with ID: {file_id} via URL: {url}")
            response = requests.delete(url, headers=headers)
            
            if response.status_code in [200, 204]: # 204 No Content is also success
                log.info(f"File {file_id} deleted successfully from OpenWebUI.")
                return True
            elif response.status_code == 404:
                log.info(f"File {file_id} not found in OpenWebUI (already deleted or never existed). Considered success for cleanup.")
                return True
            else:
                log.error(f"Failed to delete file {file_id} from OpenWebUI: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            log.error(f"Exception during OpenWebUI file deletion for ID {file_id}: {str(e)}")
            return False
            
    def get_all_active_files_mapping(self) -> Dict[str, List[str]]:
        """
        Retrieves a map of {file_id: [chat_ids]} for all files used by *active* conversations.
        Used for safety checks before deleting OpenWebUI files.
        """
        active_chat_ids = self.get_active_conversation_ids()
        all_active_files_map: Dict[str, List[str]] = {}
        
        mappings_dir = acra_config.mappings_folder
        if not os.path.exists(mappings_dir):
            log.warning(f"Mappings directory {mappings_dir} not found. Cannot build active files map.")
            return all_active_files_map

        for filename in os.listdir(mappings_dir):
            if filename.endswith("_file_mappings.json"):
                chat_id_of_mapping_file = filename.split("_file_mappings.json")[0]
                
                if chat_id_of_mapping_file in active_chat_ids: # Only process for active chats
                    mapping_file_path = os.path.join(mappings_dir, filename)
                    try:
                        with open(mapping_file_path, 'r') as f:
                            # The stored mapping is {local_abs_path: openwebui_file_id}
                            chat_specific_mappings = json.load(f)
                            for local_path, openwebui_file_id in chat_specific_mappings.items():
                                if openwebui_file_id: # Ensure file_id is not empty
                                    if openwebui_file_id not in all_active_files_map:
                                        all_active_files_map[openwebui_file_id] = []
                                    if chat_id_of_mapping_file not in all_active_files_map[openwebui_file_id]:
                                        all_active_files_map[openwebui_file_id].append(chat_id_of_mapping_file)
                    except json.JSONDecodeError:
                        log.error(f"JSON decode error for mapping file: {mapping_file_path}")
                    except Exception as e:
                        log.error(f"Error reading or processing mapping file {mapping_file_path}: {e}")
            
        log.info(f"Built active files mapping: {len(all_active_files_map)} unique OpenWebUI file IDs are used by active chats.")
        return all_active_files_map

    def delete_openwebui_files_for_chat(self, chat_id_to_clean: str) -> dict:
        """
        Deletes OpenWebUI files associated with a specific chat_id_to_clean.
        Crucially, it first checks if those files are used by *other currently active* conversations.
        """
        if not chat_id_to_clean:
            log.warning("Attempted to delete OpenWebUI files for an unspecified chat_id.")
            return {"status": "error", "message": "Chat ID not specified", "deleted_count": 0}
            
        log.info(f"Initiating OpenWebUI file deletion process for chat_id: {chat_id_to_clean}")
        
        # 1. Get a map of ALL OpenWebUI file_ids currently used by ANY active conversation.
        #    Format: {openwebui_file_id: [list_of_active_chat_ids_using_this_file]}
        all_currently_active_files_map = self.get_all_active_files_mapping()
        
        # 2. Get the file_ids specifically mapped for the chat_id_to_clean.
        mappings_dir = acra_config.mappings_folder
        mapping_file_for_chat_to_clean = os.path.join(mappings_dir, f"{chat_id_to_clean}_file_mappings.json")
        
        openwebui_file_ids_for_this_chat: List[str] = []
        if os.path.exists(mapping_file_for_chat_to_clean):
            try:
                with open(mapping_file_for_chat_to_clean, 'r') as f:
                    # Stored mapping: {local_abs_path: openwebui_file_id}
                    mappings = json.load(f)
                    openwebui_file_ids_for_this_chat = list(set(m_id for m_id in mappings.values() if m_id)) # Unique, non-empty IDs
                log.info(f"Found {len(openwebui_file_ids_for_this_chat)} OpenWebUI file IDs mapped for chat {chat_id_to_clean}: {openwebui_file_ids_for_this_chat}")
            except Exception as e:
                log.error(f"Error reading mapping file {mapping_file_for_chat_to_clean}: {e}")
                # If mapping can't be read, we can't safely determine files to delete for this chat.
                return {"status": "error", "message": f"Could not read mapping file for {chat_id_to_clean}", "deleted_count": 0}
        else:
            log.info(f"No mapping file found for chat {chat_id_to_clean}. No OpenWebUI files to delete based on its direct mappings.")
            
        deleted_count = 0
        failed_count = 0
        skipped_due_to_active_use_elsewhere = 0
        
        for file_id_to_potentially_delete in openwebui_file_ids_for_this_chat:
            is_used_by_other_active_chats = False
            if file_id_to_potentially_delete in all_currently_active_files_map:
                chats_using_this_file = all_currently_active_files_map[file_id_to_potentially_delete]
                # Is it used by any chat OTHER than the one we are currently cleaning?
                if any(active_chat_id != chat_id_to_clean for active_chat_id in chats_using_this_file):
                    is_used_by_other_active_chats = True
                    other_users = [ch_id for ch_id in chats_using_this_file if ch_id != chat_id_to_clean]
                    log.info(f"OpenWebUI file {file_id_to_potentially_delete} (from chat {chat_id_to_clean}) is ALSO used by other active chat(s): {other_users}. Skipping deletion.")
            
            if is_used_by_other_active_chats:
                skipped_due_to_active_use_elsewhere += 1
            else:
                # Safe to delete this OpenWebUI file ID
                log.info(f"Attempting deletion of OpenWebUI file {file_id_to_potentially_delete} (associated with chat {chat_id_to_clean}, not used by other active chats).")
                if self.delete_file_from_openwebui(file_id_to_potentially_delete): # No active_files_mapping passed here
                    deleted_count += 1
                else:
                    failed_count += 1
        
        result_message = (
            f"OpenWebUI file cleanup for chat {chat_id_to_clean}: "
            f"{deleted_count} deleted, "
            f"{skipped_due_to_active_use_elsewhere} skipped (used by other active chats), "
            f"{failed_count} failed."
        )
        log.info(result_message)
        
        return {
            "status": "success" if failed_count == 0 else "partial_error",
            "message": result_message,
                "deleted_count": deleted_count,
            "skipped_count": skipped_due_to_active_use_elsewhere,
                "failed_count": failed_count,
            "inspected_file_ids_from_chat_mapping": openwebui_file_ids_for_this_chat
        }

    def pipe(self, body: dict, user_message: str, model_id: str, messages: List[dict]) -> Generator[str, None, None]:
        """
        Main pipeline processing method.
        Handles commands via CommandHandler and streams LLM responses via ModelManager.
        """
        message_lower = user_message.lower()
        __event_emitter__ = body.get("__event_emitter__") 

        handled_by_confirmation, conf_response = self.command_handler.handle_confirmation(message_lower)
        if handled_by_confirmation:
            log.info(f"Request handled by confirmation: {conf_response}")
            yield f"data: {json.dumps({'choices': [{'message': {'content': conf_response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = conf_response # Update CH's last_response
            return

        # 2. Handle Specific Commands (delegated to CommandHandler)
        command_response_str = None
        if "/summarize" in message_lower:
            command_response_str = self.command_handler.handle_summarize_command(user_message) 
        elif "/structure" in message_lower:
            command_response_str = self.command_handler.handle_structure_command()
        elif "/generate" in message_lower:
            command_response_str = self.command_handler.handle_generate_command(user_message)
        elif "/clear" in message_lower:
            command_response_str = self.command_handler.handle_clear_command(user_message)
        elif "/merge" in message_lower:
            command_response_str = self.command_handler.handle_merge_command()
        elif "/regroup" in message_lower: 
            command_response_str = self.command_handler.handle_regroup_command() 
        
        if command_response_str is not None:
            log.info(f"Command handled: {command_response_str[:200]}...") # Log truncated response
            if __event_emitter__: __event_emitter__({"type": "content", "content": command_response_str})
            yield f"data: {json.dumps({'choices': [{'message': {'content': command_response_str}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = command_response_str
            return

        # 3. Default LLM interaction if no command handled
        final_user_message_for_llm = user_message
        if not user_message.strip(): # If user sends empty message
            # Show available commands if no user message (or default greeting)
            available_commands_response = self.command_handler.get_available_commands()
            log.info("Empty user message, providing available commands.")
            yield f"data: {json.dumps({'choices': [{'message': {'content': available_commands_response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = available_commands_response
            return
        
        if self.command_handler.last_response:
             final_user_message_for_llm += f"\n\n*Previous assistant response for context:* {self.command_handler.last_response}"
        
        # Always prepend the main system_prompt (which includes file analysis if any)
        full_prompt_for_llm = self.system_prompt + "\n\n" + final_user_message_for_llm
        log.info(f"Streaming LLM response for prompt (truncated): {full_prompt_for_llm[:300]}...")

        cumulative_content = ""
        try:
            yield f"data: {json.dumps({'choices': [{'delta': {'role': 'assistant'}}]})}\n\n" # Start stream

            # Use ModelManager for streaming
            for chunk_content in model_manager.stream_response(full_prompt_for_llm):
                cumulative_content += chunk_content
                if __event_emitter__:
                    __event_emitter__(({"type": "content_delta", "delta": chunk_content}))
                
                delta_res = {"choices": [{"delta": {"content": chunk_content}}]}
                yield f"data: {json.dumps(delta_res)}\n\n"

            yield f"data: {json.dumps({'choices': [{'delta': {}, 'finish_reason': 'stop'}]})}\n\n"
            yield f"data: [DONE]\n\n"

        except Exception as e:
            error_message = f"Erreur lors du streaming de la réponse LLM: {str(e)}"
            log.error(error_message, exc_info=True)
            if __event_emitter__: __event_emitter__({"type": "error", "error": error_message})
            yield f"data: {json.dumps({'error': error_message})}\n\n"
            yield f"data: [DONE]\n\n"
            return

        self.command_handler.last_response = cumulative_content # Update last response with LLM output

pipeline = Pipeline()

if __name__ == "__main__":
    pass # Placeholder for __main__ block