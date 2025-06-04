import os
import sqlite3
import shutil
import requests
from typing import List, Set

from OLLibrary.utils.log_service import get_logger, setup_logging

# Configure logging
# We need to explicitly call setup_logging in case this module is called directly
setup_logging(app_name="ACRA_Cleanup")
logger = get_logger(__name__)

# Environment variables - these will be loaded by the caller
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")
WEBUI_DB_PATH = os.getenv("WEBUI_DB_PATH", "webui.db")
OPENWEBUI_UPLOADS = os.getenv("OPENWEBUI_UPLOADS", "open-webui/uploads")
API_URL = os.getenv("API_URL", "http://localhost:5050")

# Log environment variables at module load time
logger.info(f"Cleanup service loaded with:")
logger.info(f"UPLOAD_FOLDER: {UPLOAD_FOLDER} (absolute: {os.path.abspath(UPLOAD_FOLDER)})")
logger.info(f"OUTPUT_FOLDER: {OUTPUT_FOLDER} (absolute: {os.path.abspath(OUTPUT_FOLDER)})")
logger.info(f"WEBUI_DB_PATH: {WEBUI_DB_PATH} (absolute: {os.path.abspath(WEBUI_DB_PATH)})")
logger.info(f"OPENWEBUI_UPLOADS: {OPENWEBUI_UPLOADS} (absolute: {os.path.abspath(OPENWEBUI_UPLOADS) if os.path.exists(OPENWEBUI_UPLOADS) else 'not found'})")
logger.info(f"API_URL: {API_URL}")

def reload_env_vars():
    """
    Reload environment variables - useful if this module is imported before env vars are properly set
    """
    global UPLOAD_FOLDER, OUTPUT_FOLDER, WEBUI_DB_PATH, OPENWEBUI_UPLOADS, API_URL
    
    UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
    OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")
    WEBUI_DB_PATH = os.getenv("WEBUI_DB_PATH", "webui.db")
    OPENWEBUI_UPLOADS = os.getenv("OPENWEBUI_UPLOADS", "open-webui/uploads")
    API_URL = os.getenv("API_URL", "http://localhost:5050")
    
    logger.info(f"Environment variables reloaded:")
    logger.info(f"UPLOAD_FOLDER: {UPLOAD_FOLDER} (absolute: {os.path.abspath(UPLOAD_FOLDER)})")
    logger.info(f"OUTPUT_FOLDER: {OUTPUT_FOLDER} (absolute: {os.path.abspath(OUTPUT_FOLDER)})")
    logger.info(f"WEBUI_DB_PATH: {WEBUI_DB_PATH} (absolute: {os.path.abspath(WEBUI_DB_PATH)})")
    logger.info(f"OPENWEBUI_UPLOADS: {OPENWEBUI_UPLOADS} (absolute: {os.path.abspath(OPENWEBUI_UPLOADS) if os.path.exists(OPENWEBUI_UPLOADS) else 'not found'})")
    logger.info(f"API_URL: {API_URL}")

def get_folder_ids() -> Set[str]:
    """
    Get all folder IDs from both pptx_folder and OUTPUT directories.
    
    Returns:
        Set[str]: A set of unique folder IDs
    """
    folder_ids = set()
    
    # Get folders from pptx_folder
    if os.path.exists(UPLOAD_FOLDER):
        for item in os.listdir(UPLOAD_FOLDER):
            if os.path.isdir(os.path.join(UPLOAD_FOLDER, item)):
                folder_ids.add(item)
    
    # Get folders from OUTPUT
    if os.path.exists(OUTPUT_FOLDER):
        for item in os.listdir(OUTPUT_FOLDER):
            if os.path.isdir(os.path.join(OUTPUT_FOLDER, item)):
                folder_ids.add(item)
    
    logger.debug(f"Found {len(folder_ids)} folders: {folder_ids}")
    return folder_ids

def get_chat_ids_from_db() -> Set[str]:
    """
    Get all chat IDs from the webui.db database.
    
    Returns:
        Set[str]: A set of chat IDs from the database
    """
    chat_ids = set()
    
    try:
        # Verify that the database file exists
        if not os.path.exists(WEBUI_DB_PATH):
            logger.error(f"Database file not found: {WEBUI_DB_PATH}")
            return chat_ids
            
        # Connect to the SQLite database
        conn = sqlite3.connect(WEBUI_DB_PATH)
        cursor = conn.cursor()
        
        # Log the full path to the database
        logger.info(f"Connected to database: {os.path.abspath(WEBUI_DB_PATH)}")
        
        # Verify the database schema - check if the chat table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='chat'")
        if not cursor.fetchone():
            logger.error("Chat table not found in database")
            conn.close()
            return chat_ids
        
        # Query to get all chat IDs from the chat table
        cursor.execute("SELECT id FROM chat")
        rows = cursor.fetchall()
        
        # Add each chat ID to the set
        for row in rows:
            chat_ids.add(row[0])
        
        logger.info(f"Retrieved {len(chat_ids)} chat IDs from database")
        logger.debug(f"Chat IDs: {chat_ids}")
        
        conn.close()
    except Exception as e:
        logger.error(f"Error accessing database: {str(e)}", exc_info=True)
    
    return chat_ids

def list_files_in_folder(folder_path: str) -> List[str]:
    """
    List all files in a folder.
    
    Args:
        folder_path (str): Path to the folder
        
    Returns:
        List[str]: List of filenames in the folder
    """
    if not os.path.exists(folder_path):
        logger.debug(f"Folder does not exist: {folder_path}")
        return []
    
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    logger.debug(f"Found {len(files)} files in {folder_path}")
    return files

def delete_matching_files_in_openwebui(folder_id: str):
    """
    Delete files in open-webui/uploads that match files in pptx_folder/id.
    
    Args:
        folder_id (str): The folder ID to process
        
    Returns:
        List[str]: List of deleted files
    """
    deleted_files = []
    pptx_folder_path = os.path.join(UPLOAD_FOLDER, folder_id)
    openwebui_uploads_path = OPENWEBUI_UPLOADS
    
    # Essayer de trouver le dossier uploads si le chemin par défaut ne fonctionne pas
    if not os.path.exists(openwebui_uploads_path):
        logger.warning(f"Dossier uploads non trouvé à l'emplacement {openwebui_uploads_path}")
        
        # Essayer avec un chemin absolu
        abs_path = os.path.abspath(openwebui_uploads_path)
        logger.info(f"Tentative avec le chemin absolu: {abs_path}")
        if os.path.exists(abs_path):
            openwebui_uploads_path = abs_path
            logger.info(f"Dossier uploads trouvé à: {openwebui_uploads_path}")
        else:
            # Essayer avec ./open-webui/uploads
            alternate_path = os.path.abspath("./open-webui/uploads")
            logger.info(f"Tentative avec le chemin alternatif: {alternate_path}")
            if os.path.exists(alternate_path):
                openwebui_uploads_path = alternate_path
                logger.info(f"Dossier uploads trouvé à: {openwebui_uploads_path}")
            else:
                logger.error("Impossible de trouver le dossier uploads")
    
    logger.info(f"Checking for matching files between {pptx_folder_path} and {openwebui_uploads_path}")
    
    if not os.path.exists(pptx_folder_path) or not os.path.exists(openwebui_uploads_path):
        logger.warning(f"One of the paths does not exist: {pptx_folder_path} or {openwebui_uploads_path}")
        return deleted_files
    
    # Get list of files in pptx_folder/id
    pptx_files = list_files_in_folder(pptx_folder_path)
    logger.info(f"Fichiers trouvés dans {pptx_folder_path}: {pptx_files}")
    
    # Get list of files in open-webui/uploads
    openwebui_files = list_files_in_folder(openwebui_uploads_path)
    logger.info(f"Fichiers trouvés dans {openwebui_uploads_path}: {openwebui_files}")
    
    # Delete matching files
    for pptx_file in pptx_files:
        for openwebui_file in openwebui_files:
            # Au lieu de rechercher une correspondance exacte, vérifier si le nom est inclus
            if pptx_file in openwebui_file or openwebui_file in pptx_file:  # Check if either filename is contained in the other
                file_path = os.path.join(openwebui_uploads_path, openwebui_file)
                logger.info(f"Correspondance trouvée: {pptx_file} est lié à {openwebui_file}")
                try:
                    os.remove(file_path)
                    deleted_files.append(file_path)
                    logger.info(f"Deleted file: {file_path}")
                except Exception as e:
                    logger.error(f"Error deleting file {file_path}: {str(e)}")
    
    logger.info(f"Suppression terminée: {len(deleted_files)} fichiers supprimés")
    return deleted_files

def delete_folder_and_contents(folder_path: str):
    """
    Delete a folder and all its contents.
    
    Args:
        folder_path (str): Path to the folder to delete
    """
    if not os.path.exists(folder_path):
        logger.debug(f"Folder does not exist, skipping deletion: {folder_path}")
        return
    
    try:
        shutil.rmtree(folder_path)
        logger.info(f"Deleted folder: {folder_path}")
    except Exception as e:
        logger.error(f"Error deleting folder {folder_path}: {str(e)}")

def delete_all_files_in_folder(folder_path: str):
    """
    Delete all files in a folder without deleting the folder itself.
    Reusing logic from delete_all_pptx_files in api.py.
    
    Args:
        folder_path (str): Path to the folder
    """
    if not os.path.exists(folder_path):
        logger.debug(f"Folder does not exist, skipping file deletion: {folder_path}")
        return
    
    # List files in the folder
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    if not files:
        logger.info(f"No files to delete in {folder_path}")
        return
    
    # Delete each file
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            os.remove(file_path)
            logger.info(f"Deleted file: {file_path}")
        except Exception as e:
            logger.error(f"Error deleting file {file_path}: {str(e)}")

def delete_using_api(folder_id: str):
    """
    Try to use the existing API to delete files.
    
    Args:
        folder_id (str): The folder ID to process
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        url = f"{API_URL}/delete_all_pptx_files/{folder_id}"
        logger.info(f"Calling API to delete files: {url}")
        response = requests.delete(url)
        if response.status_code == 200:
            logger.info(f"Successfully deleted files for {folder_id} using API")
            return True
        else:
            logger.error(f"API deletion failed with status code {response.status_code}: {response.text}")
            return False
    except Exception as e:
        logger.error(f"Error calling delete API: {str(e)}")
        return False

def delete_output_folder_using_api(folder_id: str):
    """
    Try to use an API to delete files in the OUTPUT folder.
    
    Args:
        folder_id (str): The folder ID to process
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Check if there's an API endpoint for deleting OUTPUT files
        # If not available, we'll use a workaround to call the existing API
        url = f"{API_URL}/delete_output_files/{folder_id}"
        logger.info(f"Calling API to delete OUTPUT files: {url}")
        response = requests.delete(url)
        if response.status_code == 200:
            logger.info(f"Successfully deleted OUTPUT files for {folder_id} using API")
            return True
        else:
            logger.error(f"API deletion for OUTPUT failed with status code {response.status_code}")
            return False
    except Exception as e:
        logger.error(f"Error calling OUTPUT delete API: {str(e)}")
        return False

def cleanup_orphaned_folder(folder_id: str):
    """
    Clean up an orphaned folder by:
    1. Deleting matching files in open-webui/uploads
    2. Using the API to delete files in pptx_folder/id
    3. Deleting the folder pptx_folder/id
    4. Using the API to delete files in OUTPUT/id (if available)
    5. Deleting all files in OUTPUT/id (if API fails)
    6. Deleting the folder OUTPUT/id
    
    Args:
        folder_id (str): The orphaned folder ID to clean up
    """
    logger.info(f"Cleaning up orphaned folder: {folder_id}")
    
    # Paths
    pptx_folder_path = os.path.join(UPLOAD_FOLDER, folder_id)
    output_folder_path = os.path.join(OUTPUT_FOLDER, folder_id)
    
    # 1. First delete files in open-webui/uploads that match files in pptx_folder/id
    deleted_files = delete_matching_files_in_openwebui(folder_id)
    logger.info(f"Deleted {len(deleted_files)} matching files from open-webui/uploads")
    
    # 2. Use the API to delete files in pptx_folder/id
    api_success = delete_using_api(folder_id)
    
    # 3. If API fails, manually delete files in pptx_folder/id
    if not api_success:
        logger.info(f"API deletion failed for {folder_id}, falling back to manual deletion")
        delete_all_files_in_folder(pptx_folder_path)
    
    # 4. Delete the folder pptx_folder/id
    delete_folder_and_contents(pptx_folder_path)
    
    # 5. Try to use API to delete files in OUTPUT/id
    output_api_success = delete_output_folder_using_api(folder_id)
    
    # 6. If API fails, manually delete files in OUTPUT/id
    if not output_api_success:
        logger.info(f"API deletion failed for OUTPUT/{folder_id}, falling back to manual deletion")
        delete_all_files_in_folder(output_folder_path)
    
    # 7. Delete the folder OUTPUT/id
    delete_folder_and_contents(output_folder_path)
    
    logger.info(f"Completed cleanup for {folder_id}")

def verify_chat_is_orphaned(folder_id: str, chat_ids: Set[str]) -> bool:
    """
    Double-check that a chat is truly orphaned by verifying against the database.
    
    Args:
        folder_id (str): The folder ID to verify
        chat_ids (Set[str]): Set of known chat IDs
        
    Returns:
        bool: True if the chat is orphaned, False if it exists in the database
    """
    # If it's already in our set of known chat IDs, it's not orphaned
    if folder_id in chat_ids:
        logger.info(f"Chat {folder_id} found in chat_ids cache - not orphaned")
        return False
    
    # Double-check directly with the database
    try:
        # Connect to the database again for a focused check
        if not os.path.exists(WEBUI_DB_PATH):
            logger.error(f"Database file not found in verification: {WEBUI_DB_PATH}")
            return False  # If we can't verify, assume it's not orphaned for safety
            
        conn = sqlite3.connect(WEBUI_DB_PATH)
        cursor = conn.cursor()
        
        # Query specifically for this chat ID
        cursor.execute("SELECT 1 FROM chat WHERE id = ?", (folder_id,))
        result = cursor.fetchone()
        
        conn.close()
        
        if result:
            logger.info(f"Chat {folder_id} found in database on verification - not orphaned")
            return False
        
        logger.info(f"Verified that chat {folder_id} is truly orphaned")
        return True
        
    except Exception as e:
        logger.error(f"Error during chat verification: {str(e)}", exc_info=True)
        return False  # If verification fails, assume it's not orphaned for safety

def cleanup_orphaned_folders():
    """
    Main function to execute the cleanup process:
    1. Get all folder IDs from pptx_folder and OUTPUT
    2. Get all chat IDs from the database
    3. Identify orphaned folders (those not in the database)
    4. Clean up each orphaned folder
    
    Returns:
        dict: Summary of cleanup operation
    """
    logger.info("Starting cleanup of orphaned folders...")
    
    # Reload environment variables to ensure we have the latest values
    reload_env_vars()
    
    # Get all folder IDs
    folder_ids = get_folder_ids()
    logger.info(f"Found {len(folder_ids)} total folders")
    
    # Get all chat IDs from the database
    chat_ids = get_chat_ids_from_db()
    logger.info(f"Found {len(chat_ids)} chats in database")
    
    # Identify potentially orphaned folders
    potential_orphaned_ids = folder_ids - chat_ids
    logger.info(f"Found {len(potential_orphaned_ids)} potentially orphaned folders")
    
    # Verify each potentially orphaned folder
    truly_orphaned_ids = set()
    for folder_id in potential_orphaned_ids:
        if verify_chat_is_orphaned(folder_id, chat_ids):
            truly_orphaned_ids.add(folder_id)
    
    logger.info(f"After verification, found {len(truly_orphaned_ids)} truly orphaned folders")
    
    # Clean up each orphaned folder
    cleaned_folders = []
    for folder_id in truly_orphaned_ids:
        cleanup_orphaned_folder(folder_id)
        cleaned_folders.append(folder_id)
    
    logger.info("Cleanup completed!")
    
    return {
        "total_folders": len(folder_ids),
        "total_chats": len(chat_ids),
        "potential_orphaned_folders": len(potential_orphaned_ids),
        "truly_orphaned_folders": len(truly_orphaned_ids),
        "cleaned_folders": cleaned_folders
    } 