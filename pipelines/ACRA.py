import json
import os
import sys
import shutil
import requests
import sqlite3
import time
import datetime
from typing import List, Generator, Dict, Any, Optional, Tuple

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
        """
        Initialize the ACRA Pipeline.
        
        This constructor:
        1. Sets up configuration and ensures required directories exist
        2. Initializes the chat_id (empty initially, set during request processing)
        3. Creates a FileManager instance for handling file operations
        4. Creates a CommandHandler instance for processing commands
        5. Validates required configuration (e.g., API keys)
        
        The pipeline is the main orchestrator for the ACRA system, handling:
        - Chat management and conversation tracking
        - File processing and PowerPoint generation
        - Command handling and LLM interactions
        """
        log.info("Initializing ACRA Pipeline")
        
        # Initialize configuration and ensure directories exist
        acra_config.ensure_directories()
        
        # State tracking - chat_id needs to be initialized before FileManager
        self.chat_id = "" 
        
        # Initialize file manager
        self.file_manager = FileManager(chat_id=self.chat_id)
        
        # Initialize command handler, passing the file_manager instance
        self.command_handler = CommandHandler(file_manager=self.file_manager)
        
        self.system_prompt = ""
        
        # Validate required configuration (OPENWEBUI_API_KEY is checked by acra_config or FileManager if needed)
        if not acra_config.get("OPENWEBUI_API_KEY"):
            log.error("OPENWEBUI_API_KEY is not set in acra_config")
            raise ValueError("OPENWEBUI_API_KEY is not set")
        
        log.info("ACRA Pipeline initialized successfully")

    def generate_report(self, info: str):
        """
        Generate a report from provided text using API or direct function call.
        Creates a new file with unique timestamp for each call.
        
        Args:
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
        """
        Reset conversation-specific state variables.
        
        This method is called when switching to a different conversation (chat_id).
        It ensures that state from a previous conversation doesn't leak into the new one.
        
        State reset includes:
        - CommandHandler state (cached structure, confirmation state, etc.)
        - System prompt used for LLM context
        """
        log.info(f"Resetting conversation state for chat_id: {self.chat_id}")
        self.command_handler.reset_state()
        self.system_prompt = ""

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

    def analyze_slide_structure(self):
        """
        Analyze the structure of PowerPoint files in the current chat's folder.
        
        This method extracts information from PowerPoint files, including projects,
        alerts, advancements, and upcoming events. It provides the core data
        that powers the system's understanding of the files' content.
        
        The method can operate in two modes:
        - API mode: Calls a remote endpoint to perform the analysis
        - Direct mode: Calls the get_slide_structure function locally
        
        Returns:
            dict: A structured representation of the PowerPoint content, including
                 projects, alerts, advancements, and upcoming events
        
        Raises:
            ValueError: If no chat_id is set
        """
        current_chat_id = self.file_manager.chat_id
        
        if not current_chat_id: # Check if chat_id is actually set
            log.error("Chat ID is not set. Cannot analyze slide structure.")
            raise ValueError("Chat ID is not set. Cannot analyze slide structure.")
        
        log.info(f"Analyzing slide structure for chat: {current_chat_id}")
        
        if acra_config.get("USE_API"):
            log.info("Using API endpoint to get slide structure")
            return self._fetch_api(f"get_slide_structure/{current_chat_id}")
        
        log.info("Using direct function call to get slide structure")
        return get_slide_structure(current_chat_id) # get_slide_structure needs chat_id
    
    def delete_all_files(self):
        """
        Deletes all files for a given chat_id.
        Manages local file deletion and OpenWebUI file deletion.
        """
        current_chat_id = self.file_manager.chat_id
        log.info(f"Deleting all files for chat: {current_chat_id}")

        # 1. Delete local files and mappings using FileManager
        fm_delete_result = self.file_manager.delete_conversation_files()
        deleted_local_files = fm_delete_result.get("deleted_count", 0)
        log.info(f"FileManager local deletion result: {fm_delete_result}")
        
        # 2. Delete OpenWebUI files associated with this chat
        webui_result = self.delete_openwebui_files_for_chat(current_chat_id) 
        deleted_webui_files = webui_result.get("deleted_count", 0)
        log.info(f"OpenWebUI deletion result: {webui_result}")
        
        # Reset pipeline state associated with files
        if self.command_handler: # Ensure command_handler is initialized
            self.command_handler.cached_structure = None

        message = f"{deleted_local_files} fichiers locaux supprimés. {deleted_webui_files} fichiers supprimés d'OpenWebUI."
        return {"message": message, "deleted_local_files": deleted_local_files, "deleted_webui_files": deleted_webui_files}

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
        """
        Process incoming requests and prepare the system for handling user messages.
        
        This method:
        1. Extracts metadata from the request body
        2. Handles chat_id management (switching between conversations)
        3. Processes uploaded files, copying them to the appropriate folders
        4. Analyzes file structure and builds the system prompt
        5. Manages cleanup of old or orphaned conversations
        
        This is called before the pipe method to set up the context for processing.
        
        Args:
            body (dict): Request body containing metadata and file information
            user (dict): User information
            
        Returns:
            dict: The processed body, potentially with added metadata
        """
        log.info(f"Received body: {body}")
        metadata = body.get("metadata", {})
        log.info(f"Metadata: {metadata}")
        
        # Get current chat_id from FileManager
        current_fm_chat_id = self.file_manager.chat_id
        log.info(f"Current state - Pipeline.chat_id: '{self.chat_id}', FileManager.chat_id: '{current_fm_chat_id}'")
        
        # Extract chat_id from the request metadata
        new_chat_id_from_request = metadata.get("chat_id")
        
        if new_chat_id_from_request:
            log.info(f"Chat_id from request metadata: '{new_chat_id_from_request}'")
            
            # Check if we're switching to a different chat
            if new_chat_id_from_request != current_fm_chat_id:
                log.info(f"*** CHAT ID CHANGING *** from FileManager current '{current_fm_chat_id}' to new '{new_chat_id_from_request}'")
                
                # FileManager's set_chat_id handles saving old mappings and loading new ones.
                self.file_manager.set_chat_id(new_chat_id_from_request)
                
                # Update Pipeline's chat_id as well, though FileManager is the primary owner now.
                self.chat_id = new_chat_id_from_request 
                
                # Reset conversation-specific states
                self.reset_conversation_state()
                
                # Clean up orphaned conversations now that active chats might have changed
                self._cleanup_old_chat(old_chat_id=current_fm_chat_id, new_chat_id=new_chat_id_from_request)

            elif not current_fm_chat_id: # First time FileManager gets a chat_id
                log.info(f"Setting initial FileManager chat_id to {new_chat_id_from_request}")
                self.file_manager.set_chat_id(new_chat_id_from_request)
                self.chat_id = new_chat_id_from_request # Keep Pipeline's chat_id in sync
        else:
            log.warning("No chat_id found in request metadata!")

        # Process files if they exist in the metadata
        files_metadata = metadata.get("files", [])
        if files_metadata:
            # Reset cached structure when new files are received
            self.command_handler.cached_structure = None
            
            # Process each uploaded file
            for file_entry in files_metadata:
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                openwebui_file_id = file_data.get("id", "N/A") # This is OpenWebUI's internal file ID

                # Construct path to the uploaded file in OpenWebUI's storage
                openwebui_uploads_dir = acra_config.get("OPENWEBUI_UPLOADS", "open-webui/uploads")
                source_path_in_openwebui_volume = os.path.join(openwebui_uploads_dir, f"{openwebui_file_id}_{filename}")
                
                # Handle case where file isn't found at expected path
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
                
                # Copy the file to the conversation's upload folder
                copied_file_path = self.file_manager.copy_uploaded_file(
                    source_path=source_path_in_openwebui_volume, 
                    filename=f"{openwebui_file_id}_{filename}"
                )
                
                # Add mapping between local file path and OpenWebUI file ID
                abs_copied_path = os.path.abspath(copied_file_path)
                if abs_copied_path not in self.file_manager.file_id_mapping:
                    self.file_manager.file_id_mapping[abs_copied_path] = openwebui_file_id
                    log.info(f"Manually added mapping for copied file: {abs_copied_path} -> {openwebui_file_id}")
                
                # Extract service name from filename
                service_name = model_manager.extract_service_name(filename)
                log.info(f"File {filename} (OpenWebUI ID: {openwebui_file_id}) identified as service: {service_name}")
            
            # Analyze structure of all files in the current chat's folder
            structure_response_data = self.analyze_slide_structure() 
            
            # Handle response from structure analysis
            if isinstance(structure_response_data, dict) and "error" in structure_response_data:
                display_response = f"Error analyzing structure: {structure_response_data['error']}"
                self.command_handler.cached_structure = {"error": display_response} # Cache error state
            elif isinstance(structure_response_data, dict):
                self.command_handler.cached_structure = structure_response_data # Cache in CommandHandler
                display_response = self.command_handler._format_slide_data(structure_response_data) # Format for system prompt
            else: # Should not happen if analyze_slide_structure is consistent
                display_response = "Error: Unexpected response type from structure analysis."
                self.command_handler.cached_structure = {"error": display_response}

            # Build system prompt with file analysis information
            self.system_prompt = f"# Here is information from the PPTX files - all information is important for understanding the user's message and the data is organized: \n\n{display_response}\n\n# Here is the user's message: "
            
            # Save file mappings for future reference
            self.file_manager.save_file_mappings()
        
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
        Main pipeline processing method that handles user messages and commands.
        
        This method:
        1. Checks if the message is a confirmation response
        2. Handles specific commands (/summarize, /structure, etc.)
        3. Falls back to LLM interaction for regular messages
        4. Streams responses back to the user
        
        Args:
            body (dict): Request body containing metadata
            user_message (str): The user's message
            model_id (str): ID of the model to use
            messages (List[dict]): Message history
            
        Yields:
            str: Streaming response chunks formatted as SSE events
        """
        message_lower = user_message.lower()
        __event_emitter__ = body.get("__event_emitter__") 

        # Step 1: Check if this is a confirmation response (yes/no)
        handled_by_confirmation, conf_response = self.command_handler.handle_confirmation(message_lower)
        if handled_by_confirmation:
            log.info(f"Request handled by confirmation: {conf_response}")
            yield f"data: {json.dumps({'choices': [{'message': {'content': conf_response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = conf_response # Update CH's last_response
            return

        # Step 2: Handle specific commands
        command_response_str = None
        if "/summarize" in message_lower:
            # Handle summarize command - generates a PowerPoint summary of uploaded files
            command_response_str = self.command_handler.handle_summarize_command(user_message) 
        elif "/structure" in message_lower:
            # Handle structure command - analyzes and displays the structure of uploaded files
            command_response_str = self.command_handler.handle_structure_command()
        elif "/generate" in message_lower:
            # Handle generate command - creates a PowerPoint from text input
            command_response_str = self.command_handler.handle_generate_command(user_message)
        elif "/clear" in message_lower:
            # Handle clear command - cleans up orphaned folders and files
            command_response_str = self.command_handler.handle_clear_command(user_message)
        elif "/merge" in message_lower:
            # Handle merge command - combines multiple PowerPoint files
            command_response_str = self.command_handler.handle_merge_command()
        elif "/regroup" in message_lower: 
            # Handle regroup command - reorganizes projects with similar topics
            command_response_str = self.command_handler.handle_regroup_command() 
        
        # If a command was handled, return its response
        if command_response_str is not None:
            log.info(f"Command handled: {command_response_str[:200]}...") # Log truncated response
            if __event_emitter__: __event_emitter__({"type": "content", "content": command_response_str})
            yield f"data: {json.dumps({'choices': [{'message': {'content': command_response_str}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = command_response_str
            return

        # Step 3: Default LLM interaction if no command handled
        final_user_message_for_llm = user_message
        if not user_message.strip(): # If user sends empty message
            # Show available commands if no user message (or default greeting)
            available_commands_response = self.command_handler.get_available_commands()
            log.info("Empty user message, providing available commands.")
            yield f"data: {json.dumps({'choices': [{'message': {'content': available_commands_response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.command_handler.last_response = available_commands_response
            return
        
        # Add the previous assistant response for context if available
        if self.command_handler.last_response:
             final_user_message_for_llm += f"\n\n*Previous assistant response for context:* {self.command_handler.last_response}"
        
        # Always prepend the main system_prompt (which includes file analysis if any)
        full_prompt_for_llm = self.system_prompt + "\n\n" + final_user_message_for_llm
        log.info(f"Streaming LLM response for prompt (truncated): {full_prompt_for_llm[:300]}...")

        # Stream the LLM response
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
            log.error(error_message)
            if __event_emitter__: __event_emitter__({"type": "error", "error": error_message})
            yield f"data: {json.dumps({'error': error_message})}\n\n"
            yield f"data: [DONE]\n\n"
            return

        self.command_handler.last_response = cumulative_content # Update last response with LLM output

pipeline = Pipeline()

if __name__ == "__main__":
    pass # Placeholder for __main__ block