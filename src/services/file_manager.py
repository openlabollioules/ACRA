"""
File Management Service for ACRA
Centralized file operations, mappings, and OpenWebUI interactions
"""
import os
import json
import shutil
import requests
from typing import Dict, List, Optional, Set, Tuple
from OLLibrary.utils.log_service import get_logger
from config_pipeline import acra_config

log = get_logger(__name__)

class FileManager:
    """
    Centralized file management for ACRA pipeline.
    Handles file operations, mappings, and OpenWebUI interactions.
    """
    
    def __init__(self, chat_id: str = None):
        self.chat_id = chat_id
        self.file_id_mapping: Dict[str, str] = {}  # {file_path: file_id}
        
        # Ensure directories exist
        acra_config.ensure_directories()
        
        if chat_id:
            self.load_file_mappings()
    
    def set_chat_id(self, chat_id: str):
        """Set or change the chat ID and load corresponding mappings"""
        if self.chat_id != chat_id:
            if self.chat_id:
                self.save_file_mappings()  # Save current mappings
            
            self.chat_id = chat_id
            self.file_id_mapping = {}
            self.load_file_mappings()
            
            # Ensure conversation directories exist
            os.makedirs(acra_config.get_conversation_upload_folder(chat_id), exist_ok=True)
            os.makedirs(acra_config.get_conversation_output_folder(chat_id), exist_ok=True)
    
    def save_file_mappings(self):
        """Save file mappings to JSON file"""
        if not self.chat_id:
            log.warning("No chat_id set, cannot save mappings")
            return
        
        try:
            mapping_file = acra_config.get_mapping_file_path(self.chat_id)
            os.makedirs(os.path.dirname(mapping_file), exist_ok=True)
            
            # Convert absolute paths to relative for portability
            relative_mappings = {}
            for file_path, file_id in self.file_id_mapping.items():
                relative_path = os.path.relpath(file_path, os.getcwd())
                relative_mappings[relative_path] = file_id
            
            with open(mapping_file, 'w') as f:
                json.dump(relative_mappings, f, indent=2)
            
            log.info(f"Saved {len(relative_mappings)} file mappings to {mapping_file}")
        except Exception as e:
            log.error(f"Error saving file mappings: {str(e)}")
    
    def load_file_mappings(self):
        """Load file mappings from JSON file"""
        if not self.chat_id:
            log.warning("No chat_id set, cannot load mappings")
            return
        
        try:
            mapping_file = acra_config.get_mapping_file_path(self.chat_id)
            
            if os.path.exists(mapping_file):
                with open(mapping_file, 'r') as f:
                    relative_mappings = json.load(f)
                
                # Convert relative paths back to absolute
                self.file_id_mapping = {}
                for relative_path, file_id in relative_mappings.items():
                    abs_path = os.path.abspath(os.path.join(os.getcwd(), relative_path))
                    self.file_id_mapping[abs_path] = file_id
                
                log.info(f"Loaded {len(self.file_id_mapping)} file mappings")
            else:
                log.info(f"No mapping file found for chat {self.chat_id}")
                self.file_id_mapping = {}
        except Exception as e:
            log.error(f"Error loading file mappings: {str(e)}")
            self.file_id_mapping = {}
    
    def upload_to_openwebui(self, file_path: str) -> Dict[str, str]:
        """
        Upload a file to OpenWebUI and return download URL.
        Uses mapping to avoid duplicate uploads.
        """
        try:
            abs_file_path = os.path.abspath(file_path)
            
            # Check if file already uploaded
            if abs_file_path in self.file_id_mapping:
                file_id = self.file_id_mapping[abs_file_path]
                log.info(f"File already uploaded, reusing ID: {file_id}")
                download_url = f"http://localhost:3030/api/v1/files/{file_id}/content"
                return {"download_url": download_url}
            
            # Upload new file
            headers = {
                "accept": "application/json",
                "Authorization": f"Bearer {acra_config.get('OPENWEBUI_API_KEY')}"
            }
            
            url = f"{acra_config.get('OPENWEBUI_API_URL')}files/"
            log.info(f"Uploading file to OpenWebUI: {file_path}")
            
            with open(file_path, "rb") as f:
                files = {"file": (os.path.basename(file_path), f, "application/octet-stream")}
                response = requests.post(url, headers=headers, files=files)
            
            if response.status_code != 200:
                log.error(f"File upload failed: {response.status_code} - {response.text}")
                return {"error": f"File upload failed: {response.status_code}"}
            
            response_data = response.json()
            file_id = response_data.get("id", "")
            
            if not file_id:
                log.error("No file ID returned from upload")
                return {"error": "No file ID returned from upload"}
            
            # Store in mapping
            self.file_id_mapping[abs_file_path] = file_id
            log.info(f"Added file mapping: {abs_file_path} -> {file_id}")
            
            download_url = f"http://localhost:3030/api/v1/files/{file_id}/content"
            return {"download_url": download_url}
            
        except Exception as e:
            log.error(f"Error uploading file to OpenWebUI: {str(e)}")
            return {"error": f"Error uploading file: {str(e)}"}
    
    def copy_uploaded_file(self, source_path: str, filename: str) -> str:
        """Copy an uploaded file to the conversation folder"""
        if not self.chat_id:
            raise ValueError("No chat_id set")
        
        destination_folder = acra_config.get_conversation_upload_folder(self.chat_id)
        destination_path = os.path.join(destination_folder, filename)
        
        shutil.copy(source_path, destination_path)
        log.info(f"Copied file from {source_path} to {destination_path}")
        
        return destination_path
    
    def get_conversation_files(self, file_extensions: List[str] = None) -> List[str]:
        """Get list of files in the conversation folder"""
        if not self.chat_id:
            return []
        
        if file_extensions is None:
            file_extensions = acra_config.get("ALLOWED_EXTENSIONS", [".pptx"])
        
        folder_path = acra_config.get_conversation_upload_folder(self.chat_id)
        if not os.path.exists(folder_path):
            return []
        
        files = []
        for filename in os.listdir(folder_path):
            if any(filename.lower().endswith(ext) for ext in file_extensions):
                files.append(os.path.join(folder_path, filename))
        
        return files
    
    def delete_file(self, file_path: str, update_mapping: bool = True) -> Dict[str, str]:
        """Delete a file and optionally update mapping"""
        try:
            abs_path = os.path.abspath(file_path)
            
            if not os.path.exists(abs_path):
                return {"error": "File not found"}
            
            os.remove(abs_path)
            log.info(f"Deleted file: {abs_path}")
            
            if update_mapping and abs_path in self.file_id_mapping:
                del self.file_id_mapping[abs_path]
                log.info(f"Removed mapping for deleted file: {abs_path}")
                self.save_file_mappings()
            
            return {"message": f"File {os.path.basename(file_path)} deleted successfully"}
        except Exception as e:
            log.error(f"Error deleting file {file_path}: {str(e)}")
            return {"error": f"Error deleting file: {str(e)}"}
    
    def delete_conversation_files(self) -> Dict[str, any]:
        """Delete all files for the current conversation"""
        if not self.chat_id:
            return {"error": "No chat_id set"}
        
        deleted_count = 0
        
        # Delete files from upload folder
        upload_folder = acra_config.get_conversation_upload_folder(self.chat_id)
        if os.path.exists(upload_folder):
            for filename in os.listdir(upload_folder):
                file_path = os.path.join(upload_folder, filename)
                if os.path.isfile(file_path):
                    result = self.delete_file(file_path, update_mapping=False)
                    if "error" not in result:
                        deleted_count += 1
        
        # Delete files from output folder
        output_folder = acra_config.get_conversation_output_folder(self.chat_id)
        if os.path.exists(output_folder):
            for filename in os.listdir(output_folder):
                if filename.endswith(".pptx"):
                    file_path = os.path.join(output_folder, filename)
                    if os.path.isfile(file_path):
                        result = self.delete_file(file_path, update_mapping=False)
                        if "error" not in result:
                            deleted_count += 1
        
        # Clear mappings and delete mapping file
        self.file_id_mapping = {}
        mapping_file = acra_config.get_mapping_file_path(self.chat_id)
        if os.path.exists(mapping_file):
            os.remove(mapping_file)
            log.info(f"Deleted mapping file: {mapping_file}")
        
        return {"message": f"Deleted {deleted_count} files", "deleted_count": deleted_count}
    
    def get_existing_summaries(self) -> List[Tuple[str, str]]:
        """Get list of existing summary files with their download URLs"""
        if not self.chat_id:
            return []
        
        summaries = []
        
        # Check OUTPUT folder
        output_folder = acra_config.get_conversation_output_folder(self.chat_id)
        if os.path.exists(output_folder):
            for filename in os.listdir(output_folder):
                if filename.endswith(".pptx"):
                    file_path = os.path.join(output_folder, filename)
                    upload_result = self.upload_to_openwebui(file_path)
                    
                    if "download_url" in upload_result:
                        summaries.append(("OUTPUT/" + filename, upload_result["download_url"]))
        
        # Check upload folder for summary files
        upload_folder = acra_config.get_conversation_upload_folder(self.chat_id)
        if os.path.exists(upload_folder):
            for filename in os.listdir(upload_folder):
                if filename.endswith(".pptx") and ("_summary_" in filename or "_text_summary_" in filename):
                    file_path = os.path.join(upload_folder, filename)
                    upload_result = self.upload_to_openwebui(file_path)
                    
                    if "download_url" in upload_result:
                        summaries.append(("pptx_folder/" + filename, upload_result["download_url"]))
        
        if summaries:
            self.save_file_mappings()
        
        return summaries
    
    def cleanup_orphaned_mappings(self) -> int:
        """Clean up mappings that point to non-existent files"""
        orphaned_mappings = []
        
        for file_path, file_id in list(self.file_id_mapping.items()):
            if not os.path.exists(file_path):
                orphaned_mappings.append((file_path, file_id))
        
        removed_count = 0
        for file_path, file_id in orphaned_mappings:
            log.info(f"Removing orphaned mapping: {file_path} -> {file_id}")
            del self.file_id_mapping[file_path]
            removed_count += 1
        
        if removed_count > 0:
            self.save_file_mappings()
            log.info(f"Removed {removed_count} orphaned mappings")
        
        return removed_count 