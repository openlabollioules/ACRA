import json
import os
import sys
import shutil
import requests
import uuid
import sqlite3
from typing import List, Union, Generator, Iterator, Dict, Any
from langchain_ollama import  OllamaLLM
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..","src")))
from core import summarize_ppt, get_slide_structure, delete_all_pptx_files, generate_pptx_from_text
from services import merge_pptx_files

from OLLibrary.utils.text_service import remove_tags_keep
from OLLibrary.utils.log_service import setup_logging, get_logger

import logging

# Set up the main application logger
setup_logging(app_name="ACRA_Pipeline")
# Use a specific logger for this module
log = get_logger(__name__)
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")
MAPPINGS_FOLDER = os.getenv("MAPPINGS_FOLDER", "mappings")

class Pipeline:
    def __init__(self):
        load_dotenv()
        log.info("Initializing ACRA Pipeline")
        self.last_response = None

        # Fix: Properly convert USE_API to boolean
        use_api_env = os.getenv("USE_API", "False")
        self.use_api = use_api_env.lower() in ("true", "1", "t", "yes", "y")
        print(f"USE_API: {self.use_api}")
        # self.model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", num_ctx=32000)
        self.streaming_model = OllamaLLM(model="deepseek-r1:14b", base_url="http://host.docker.internal:11434", num_ctx=131000, stream=True)

        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030/api/v1/"
        self.openwebui_db_path = os.getenv("OPENWEBUI_DB_PATH", "./open-webui/webui.db")

        self.small_model = OllamaLLM(model="qwen2.5:3b", base_url="http://host.docker.internal:11434", num_ctx=131000)

        self.file_path_list = []
        self.openwebui_api_key = os.getenv("OPENWEBUI_API_KEY")
        if not self.openwebui_api_key:
            log.error("OPENWEBUI_API_KEY is not set")
            raise ValueError("OPENWEBUI_API_KEY is not set")

        self.chat_id = ""
        self.current_chat_id = ""  # To track conversation changes
        self.small_model = OllamaLLM(model="gemma3:latest", base_url="http://host.docker.internal:11434", num_ctx=64000)
        self.system_prompt = ""
        self.message_id = 0
        
        # Variable pour stocker la structure traitée
        self.cached_structure = None

        # State tracking
        self.waiting_for_confirmation = False
        self.confirmation_command = ""
        self.confirmation_additional_info = ""
        
        # File ID mapping
        self.file_id_mapping = {}  # Maps {file_path: file_id}
        
        # Create mappings folder if it doesn't exist
        os.makedirs(MAPPINGS_FOLDER, exist_ok=True)
        
        log.info("ACRA Pipeline initialized successfully")

    def generate_report(self, foldername, info):
        """
        Génère un rapport à partir du texte fourni en utilisant une requête POST.
        Génère un nouveau fichier avec un timestamp unique à chaque appel.
        
        Args:
            foldername (str): Le nom du dossier où stocker le rapport
            info (str): Le texte à analyser pour générer le rapport
            
        Returns:
            dict: Résultat de la requête avec l'URL de téléchargement
        """
        log.info(f"Generating report for folder: {foldername}")
        log.info(f"use_api setting is: {self.use_api}")
        
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log.info(f"Creating report with timestamp: {timestamp}")
        
        if self.use_api:
            log.info("Using API endpoint to generate report")
            endpoint = f"generate_report/{foldername}?info={info}&timestamp={timestamp}"
            result = self.fetch(endpoint)
            if "error" in result:
                return result
                
            return self.download_file_openwebui(result["summary"])

        log.info("Using direct function call to generate report")
        result = generate_pptx_from_text(foldername, info, timestamp)
        if "error" in result:
            return result
            
        upload_result = self.download_file_openwebui(result["summary"])
        # Save the mapping after uploading the new file
        self.save_file_mappings()
        return upload_result

    def reset_conversation_state(self):
        """Réinitialise les états spécifiques à une conversation"""
        log.info(f"Resetting conversation state for chat_id: {self.chat_id}")
        self.last_response = None
        self.system_prompt = ""
        self.file_path_list = []
        self.message_id = 0
        self.waiting_for_confirmation = False
        self.confirmation_command = ""
        self.confirmation_additional_info = ""
        self.cached_structure = None
        # Note: We don't clear file_id_mapping here anymore
        # as we want to preserve the mapping between files and OpenWebUI file IDs
        # This mapping is loaded/saved per conversation ID

    def save_file_mappings(self):
        """
        Sauvegarde le mapping des fichiers dans un fichier JSON dans le dossier de mappings.
        """
        try:
            os.makedirs(MAPPINGS_FOLDER, exist_ok=True)
            mapping_file = os.path.join(MAPPINGS_FOLDER, f"{self.chat_id}_file_mappings.json")
            
            # Convert absolute paths to relative for better portability
            relative_mappings = {}
            for file_path, file_id in self.file_id_mapping.items():
                relative_path = os.path.relpath(file_path, os.getcwd())
                relative_mappings[relative_path] = file_id
            
            with open(mapping_file, 'w') as f:
                json.dump(relative_mappings, f)
            
            log.info(f"Saved file mappings to {mapping_file}")
        except Exception as e:
            log.error(f"Error saving file mappings: {str(e)}")

    def load_file_mappings(self):
        """
        Charge le mapping des fichiers depuis un fichier JSON dans le dossier de mappings.
        """
        try:
            mapping_file = os.path.join(MAPPINGS_FOLDER, f"{self.chat_id}_file_mappings.json")
            
            if os.path.exists(mapping_file):
                with open(mapping_file, 'r') as f:
                    relative_mappings = json.load(f)
                
                # Convert relative paths back to absolute
                self.file_id_mapping = {}
                for relative_path, file_id in relative_mappings.items():
                    abs_path = os.path.abspath(os.path.join(os.getcwd(), relative_path))
                    self.file_id_mapping[abs_path] = file_id
                
                log.info(f"Loaded {len(self.file_id_mapping)} file mappings from {mapping_file}")
            else:
                log.info(f"No mapping file found at {mapping_file}")
        except Exception as e:
            log.error(f"Error loading file mappings: {str(e)}")
            self.file_id_mapping = {}

    def fetch(self, endpoint):
            """Effectue une requête GET synchrone"""
            url = f"{self.api_url}/{endpoint}"
            log.debug(f"Fetching from: {url}")
            response = requests.get(url)
            if response.status_code != 200:
                log.error(f"API request failed: {response.status_code} - {response.text}")
            return response.json() if response.status_code == 200 else {"error": "Request failed"}

    def post(self, endpoint, data=None, files=None, headers=None):
        """Effectue une requête POST synchrone"""
        # Si l'endpoint commence par http, on le considère comme une URL complète
        print(f"Endpoint: {endpoint}")
        if endpoint.startswith("http"):
            url = endpoint
        else:
            # Sinon on le préfixe avec l'URL de l'API
            url = f"{self.api_url}/{endpoint}"
        log.debug(f"Posting to: {url}")
        response = requests.post(url, data=data, files=files, headers=headers)
        if response.status_code != 200:
            log.error(f"API POST request failed: {response.status_code} - {response.text}")
        return response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}

    # Method to download a filde using the openwebui api. 
    def download_file_openwebui(self, file: str):
        """
        Télécharge un fichier à partir d'un nom de fichier.
        Utilise un système de mapping pour éviter les téléchargements redondants.
        """
        try:
            # Check if we already have this file mapped
            file_path = os.path.abspath(file)
            if file_path in self.file_id_mapping:
                file_id = self.file_id_mapping[file_path]
                log.info(f"File already uploaded, reusing ID: {file_id}")
                download_url = f"http://localhost:3030/api/v1/files/{file_id}/content"
                log.info(f"Reusing existing download URL: {download_url}")
                return {"download_url": download_url}
            
            headers = {
                "accept": "application/json",
                # Remove Content-Type header to let requests set it automatically for multipart/form-data
                "Authorization": f"Bearer {self.openwebui_api_key}"
            }
            url = f"{self.openwebui_api}files/"  # Remove trailing slash
            log.info(f"Uploading file to OpenWebUI API: {url}")
            log.info(f"File path: {file}")
            
            file_id = ""
            try:
                with open(file, "rb") as f:
                    files = {"file": (os.path.basename(file), f, "application/octet-stream")}
                    
                    # Use direct requests instead of self.post for more control
                    response = requests.post(url, headers=headers, files=files)
                    log.info(f"Upload response status: {response.status_code}")
                    
                    if response.status_code != 200:
                        log.error(f"File upload failed: {response.status_code} - {response.text}")
                        return {"error": f"File upload failed: {response.status_code}"}
                    
                    # Parse the response
                    response_data = response.json()
                    log.info(f"Upload response: {response_data}")
                    file_id = response_data.get("id", "")
                    
                    if not file_id:
                        log.error("No file ID returned from upload")
                        return {"error": "No file ID returned from upload"}
                    
                    # Store in our mapping
                    self.file_id_mapping[file_path] = file_id
                    log.info(f"Added file mapping: {file_path} -> {file_id}")
            except Exception as e:
                log.error(f"Error uploading file: {str(e)}")
                return {"error": f"Error uploading file: {str(e)}"}
            
            download_url = f"http://localhost:3030/api/v1/files/{file_id}/content"
            log.info(f"Download URL: {download_url}")
            return {"download_url": download_url}
        except Exception as e:
            log.error(f"Error in download_file_openwebui: {str(e)}")
            return {"error": f"Error in download_file_openwebui: {str(e)}"}
    
    def summarize_folder(self, foldername=None, add_info=None):
        """
        Envoie une demande pour résumer tous les fichiers PowerPoint dans un dossier.
        Génère un nouveau fichier avec un timestamp unique à chaque appel.
        
        Args:
            foldername (str, optional): Le nom du dossier à résumer. Si None, utilise le chat_id.
            add_info (str, optional): Informations supplémentaires à ajouter au résumé.
        Returns:
            dict: Les résultats de l'opération de résumé.
        """
        if foldername is None:
            foldername = self.chat_id
        
        log.info(f"Summarizing folder: {foldername}")
        log.info(f"use_api setting is: {self.use_api}")
        log.info(f"Additional info: {add_info}")
        
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log.info(f"Creating summary with timestamp: {timestamp}")
        
        if self.use_api:
            log.info("Using API endpoint to summarize folder")
            endpoint = f"acra/{foldername}"
            if add_info:
                endpoint += f"?add_info={add_info}&timestamp={timestamp}"
            else:
                endpoint += f"?timestamp={timestamp}"
            result = self.fetch(endpoint)
            if "error" in result:
                return result
            
            # Let's make sure both the output file and the source files are properly tracked
            # First, track source files from the pptx_folder
            source_folder = os.path.join(UPLOAD_FOLDER, foldername)
            if os.path.exists(source_folder):
                for filename in os.listdir(source_folder):
                    if filename.lower().endswith(".pptx"):
                        source_file_path = os.path.join(source_folder, filename)
                        abs_file_path = os.path.abspath(source_file_path)
                        if abs_file_path not in self.file_id_mapping:
                            log.info(f"Adding source file to tracking: {abs_file_path}")
                            # Will be uploaded and added to mapping when needed
            
            # Then get the download URL for the new summary file
            upload_result = self.download_file_openwebui(result["summary"])
            self.save_file_mappings()  # Save all mappings
            return upload_result
        
        log.info("Using direct function call to summarize folder")
        result = summarize_ppt(foldername, add_info, timestamp)
        if "error" in result:
            return result
        
        # Track both the output file and source files
        upload_result = self.download_file_openwebui(result["summary"])
        
        # Track source files
        source_folder = os.path.join(UPLOAD_FOLDER, foldername)
        if os.path.exists(source_folder):
            for filename in os.listdir(source_folder):
                if filename.lower().endswith(".pptx"):
                    source_file_path = os.path.join(source_folder, filename)
                    abs_file_path = os.path.abspath(source_file_path)
                    if abs_file_path not in self.file_id_mapping:
                        log.info(f"Source file not yet tracked: {abs_file_path}")
                        # Will be uploaded and tracked when needed
        
        # Save the mapping after uploading all files
        self.save_file_mappings()
        return upload_result

    def extract_service_name(self, filename):
        """
        Extrait le nom du service à partir du nom du fichier PowerPoint en utilisant le modèle small_model.
        
        Args:
            filename (str): Le nom du fichier PowerPoint
            
        Returns:
            str: Le nom du service extrait
        """
        prompt = f"Tu es un assistant spécialisé dans le traitement automatique des noms de fichiers. On te donne un nom de fichier de présentation (PowerPoint) contenant un identifiant unique suivi du titre du document. Ton objectif est d'extraire uniquement le titre du document dans un format propre et lisible pour un humain. Le titre est toujours situé après le dernier underscore (`_`) ou après une chaîne d'identifiants. Supprime l'extension `.pptx` ou toute autre extension. Remplace les underscores (`_`) ou tirets (`-`) par des espaces, et capitalise correctement chaque mot. Exemple : **Nom de fichier :** `dc56be63-37a6-4ed6-9223-50f545028ab4_CRA_SERVICE_UX.pptx`   **Titre extrait :** `Service UX` Donne uniquement le titre extrait (pas d'explication), en une seule ligne. voici le nom du fichier : {filename}"
        
        service_name = self.small_model.invoke(prompt)
        # Nettoyer la réponse (enlever les espaces, retours à la ligne, etc.)
        return service_name.strip()

    def analyze_slide_structure(self, foldername=None):
        """
        Analyse la structure des diapositives dans un dossier.
        
        Args:
            foldername (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'analyse.
        """
        if foldername is None:
            foldername = self.chat_id
        
        log.info(f"Analyzing slide structure for folder: {foldername}")
        log.info(f"use_api setting is: {self.use_api}")
        
        if self.use_api:
            log.info("Using API endpoint to get slide structure")
            return self.fetch(f"get_slide_structure/{foldername}")
        
        log.info("Using direct function call to get slide structure")
        return get_slide_structure(foldername)
    
    def format_all_slide_data(self, data: dict) -> str:
        """
        Formate les données de plusieurs présentations PPTX en une seule chaîne de texte structurée,
        regroupant tous les projets sans séparation par fichier et avec les événements à venir par service.
        
        Si une structure traitée existe déjà en cache et que data n'est pas None, utilise la structure en cache.
        Sinon, traite la structure et la stocke en cache.

        Args:
            data (dict): Dictionnaire contenant les projets et métadonnées conforme au nouveau format.

        Returns:
            str: Une chaîne de texte structurée listant les informations de tous les projets.
        """
        # Si data est None ou vide, renvoyer un message d'erreur
        if not data:
            return "Aucun fichier PPTX fourni."
            
        # Si data est fourni, mettre à jour le cache
        self.cached_structure = data
        
        # Utiliser la structure en cache si elle existe
        structure_to_process = self.cached_structure
        
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


    def delete_all_files(self, foldername=None):
        """
        Supprime tous les fichiers dans un dossier et met à jour le système de mapping de fichiers.
        
        Args:
            foldername (str, optional): Le nom du dossier à vider. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'opération de suppression.
        """
        if foldername is None:
            foldername = self.chat_id
        
        log.info(f"Deleting all files for folder: {foldername}")
        
        # Get folder paths for both pptx_folder and OUTPUT
        pptx_folder_path = os.path.join(UPLOAD_FOLDER, foldername)
        output_folder_path = os.path.join(OUTPUT_FOLDER, foldername)
        
        deleted_count = 0
        removed_mappings = 0
        
        # Delete files from pptx_folder
        if self.use_api:
            log.info(f"Using API to delete files from {pptx_folder_path}")
            url = f"{self.api_url}/delete_all_pptx_files/{foldername}"
            response = requests.delete(url)
            result = response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}
        else:
            log.info(f"Directly deleting files from {pptx_folder_path}")
            try:
                if os.path.exists(pptx_folder_path):
                    files = os.listdir(pptx_folder_path)
                    for file in files:
                        file_path = os.path.join(pptx_folder_path, file)
                        abs_path = os.path.abspath(file_path)
                        
                        # Remove mapping if exists
                        if abs_path in self.file_id_mapping:
                            del self.file_id_mapping[abs_path]
                            removed_mappings += 1
                            log.info(f"Removed mapping for deleted file: {abs_path}")
                        
                        # Delete the file
                        os.remove(file_path)
                        deleted_count += 1
                        log.info(f"Deleted file: {file_path}")
                        
                    log.info(f"Deleted {deleted_count} files from pptx_folder")
                    result = {"message": f"{deleted_count} fichiers supprimés avec succès."}
                else:
                    result = {"message": "Le dossier n'existe pas."}
            except Exception as e:
                log.error(f"Error deleting files from pptx_folder: {str(e)}")
                result = {"error": f"Erreur lors de la suppression des fichiers: {str(e)}"}
        
        # Also clean up any files in the OUTPUT folder for this conversation
        try:
            if os.path.exists(output_folder_path):
                output_files = os.listdir(output_folder_path)
                output_deleted = 0
                
                for file in output_files:
                    if file.endswith(".pptx"):  # Only delete PPTX files, leave mapping files
                        file_path = os.path.join(output_folder_path, file)
                        abs_path = os.path.abspath(file_path)
                        
                        # Remove mapping if exists
                        if abs_path in self.file_id_mapping:
                            del self.file_id_mapping[abs_path]
                            removed_mappings += 1
                            log.info(f"Removed mapping for deleted output file: {abs_path}")
                        
                        # Delete the file
                        os.remove(file_path)
                        output_deleted += 1
                        log.info(f"Deleted output file: {file_path}")
                
                log.info(f"Deleted {output_deleted} files from OUTPUT folder")
                
                # Update the result message
                if "message" in result:
                    result["message"] += f" Plus {output_deleted} fichiers supprimés du dossier de sortie."
        except Exception as e:
            log.error(f"Error cleaning up OUTPUT folder: {str(e)}")
            # Don't override the main result if there was an error here
        
        # Save the updated mappings
        if removed_mappings > 0:
            log.info(f"Removed {removed_mappings} file mappings. Saving updated mappings.")
            self.save_file_mappings()
        
        # Reset file path list and cached structure
        self.file_path_list = []
        self.cached_structure = None
        
        return result
    
    def get_files_in_folder(self, foldername=None):
        """
        Récupère la liste des fichiers dans un dossier.
        
        Args:
            foldername (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            list: Liste des noms de fichiers PPTX dans le dossier.
        """
        if foldername is None:
            foldername = self.chat_id
            
        folder_path = os.path.join("./pptx_folder", foldername)
        if not os.path.exists(folder_path):
            return []
            
        return [f for f in os.listdir(folder_path) if f.lower().endswith(".pptx")]

    def get_active_conversation_ids(self):
        """
        Récupère tous les IDs de conversation actifs depuis la base de données OpenWebUI.
        
        Returns:
            list: Liste des IDs de conversation actifs.
        """
        conversation_ids = []
        try:
            if not os.path.exists(self.openwebui_db_path):
                log.error(f"OpenWebUI database not found at {self.openwebui_db_path}")
                return conversation_ids
                
            conn = sqlite3.connect(self.openwebui_db_path)
            cursor = conn.cursor()
            
            # Query to get all active conversation IDs
            cursor.execute("SELECT id FROM conversations WHERE deleted_at IS NULL")
            rows = cursor.fetchall()
            
            conversation_ids = [row[0] for row in rows]
            conn.close()
            
            log.info(f"Found {len(conversation_ids)} active conversations in OpenWebUI database")
        except Exception as e:
            log.error(f"Error retrieving conversation IDs from database: {str(e)}")
        
        return conversation_ids

    def cleanup_orphaned_conversations(self):
        """
        Nettoie les dossiers et fichiers de conversations qui n'existent plus dans OpenWebUI.
        Appelé quand le chat_id change pour s'assurer que les ressources sont bien gérées.
        
        Returns:
            dict: Résultats de l'opération de nettoyage
        """
        log.info("Starting cleanup of orphaned conversations")
        active_conversations = self.get_active_conversation_ids()
        
        if not active_conversations:
            log.warning("No active conversations found or unable to get conversation list")
            return {"status": "warning", "message": "Could not retrieve active conversations"}
        
        deleted_folders = 0
        deleted_files = 0
        deleted_mappings = 0
        
        # Clean pptx_folder
        try:
            if os.path.exists(UPLOAD_FOLDER):
                for folder_name in os.listdir(UPLOAD_FOLDER):
                    folder_path = os.path.join(UPLOAD_FOLDER, folder_name)
                    if os.path.isdir(folder_path) and folder_name not in active_conversations:
                        log.info(f"Deleting orphaned PPTX folder: {folder_path}")
                        # Count files before deletion
                        deleted_files += len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
                        shutil.rmtree(folder_path)
                        deleted_folders += 1
        except Exception as e:
            log.error(f"Error cleaning up pptx folders: {str(e)}")
        
        # Clean OUTPUT folder
        try:
            if os.path.exists(OUTPUT_FOLDER):
                for folder_name in os.listdir(OUTPUT_FOLDER):
                    folder_path = os.path.join(OUTPUT_FOLDER, folder_name)
                    if os.path.isdir(folder_path) and folder_name not in active_conversations:
                        log.info(f"Deleting orphaned OUTPUT folder: {folder_path}")
                        # Count files before deletion
                        deleted_files += len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
                        shutil.rmtree(folder_path)
                        deleted_folders += 1
        except Exception as e:
            log.error(f"Error cleaning up output folders: {str(e)}")
        
        # Clean mappings folder
        try:
            if os.path.exists(MAPPINGS_FOLDER):
                for filename in os.listdir(MAPPINGS_FOLDER):
                    if filename.endswith("_file_mappings.json"):
                        # Extract chat_id from filename (chat_id_file_mappings.json)
                        chat_id = filename.split("_file_mappings.json")[0]
                        if chat_id not in active_conversations:
                            mapping_path = os.path.join(MAPPINGS_FOLDER, filename)
                            log.info(f"Deleting orphaned mapping file: {mapping_path}")
                            os.remove(mapping_path)
                            deleted_mappings += 1
        except Exception as e:
            log.error(f"Error cleaning up mapping files: {str(e)}")
        
        # Clean uploads folder
        try:
            uploads_folder = "./open-webui/uploads"
            if os.path.exists(uploads_folder):
                for filename in os.listdir(uploads_folder):
                    file_path = os.path.join(uploads_folder, filename)
                    # Files in uploads are prefixed with file_id, we can't directly map to chat_id
                    # We clean these via OpenWebUI API or db directly in a future iteration
                    # For now, we keep files in uploads as they might be referenced elsewhere
        except Exception as e:
            log.error(f"Error examining uploads folder: {str(e)}")
        
        result = {
            "status": "success",
            "deleted_folders": deleted_folders,
            "deleted_files": deleted_files,
            "deleted_mappings": deleted_mappings,
            "active_conversations": len(active_conversations)
        }
        
        log.info(f"Cleanup results: {result}")
        return result

    async def inlet(self, body: dict, user: dict) -> dict:
        log.info(f"Received body: {body}")
        
        # Debug log the current state
        log.info(f"Current state - self.chat_id: '{self.chat_id}', self.current_chat_id: '{self.current_chat_id}'")
        
        # Get conversation ID from body
        if body.get("metadata", {}).get("chat_id") != None:
            new_chat_id = body.get("metadata", {}).get("chat_id", "default")
            
            # If chat_id changed, we need to save current mappings and load new ones
            if self.chat_id != new_chat_id and self.chat_id:
                log.info(f"Chat ID changed from {self.chat_id} to {new_chat_id}")
                self.save_file_mappings()  # Save mappings for old chat
                
                # Run cleanup process to check for orphaned conversations
                cleanup_result = self.cleanup_orphaned_conversations()
                log.info(f"Cleanup results: {cleanup_result}")
                
                # Update chat_id after cleanup
                self.chat_id = new_chat_id
                
                # Reset state but preserve file mappings
                self.reset_conversation_state()
                
                # Load mappings for new conversation
                self.load_file_mappings()
            elif not self.chat_id:
                # First time setting chat_id
                self.chat_id = new_chat_id
                self.load_file_mappings()  # Try to load any existing mappings
                
            if not self.current_chat_id:
                self.current_chat_id = self.chat_id

        # Create folders for the conversation
        # Create pptx_folder for source files
        conversation_folder = os.path.join(UPLOAD_FOLDER, self.chat_id)
        os.makedirs(conversation_folder, exist_ok=True)
        log.info(f"Created/verified upload folder at: {conversation_folder}")
        
        # Create output folder for generated files
        output_folder = os.path.join(OUTPUT_FOLDER, self.chat_id) 
        os.makedirs(output_folder, exist_ok=True)
        log.info(f"Created/verified output folder at: {output_folder}")
        
        # Create mappings folder if needed
        os.makedirs(MAPPINGS_FOLDER, exist_ok=True)

        # Extract files from body['metadata']['files']
        files = body.get("metadata", {}).get("files", [])
        if files:
            # Réinitialiser la structure en cache car de nouveaux fichiers ont été ajoutés
            self.cached_structure = None
            
            # Traiter les fichiers
            for file_entry in files:
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                file_id = file_data.get("id", "N/A")

                filecomplete_name = file_id + "_" + filename

                source_path = os.path.join("./uploads", filecomplete_name)
                # Update destination to use conversation foldername
                destination_path = os.path.join(conversation_folder, filecomplete_name)
                
                self.file_path_list.append(destination_path)
                shutil.copy(source_path, destination_path)
                
                # Add to file mappings if not already present
                abs_path = os.path.abspath(destination_path)
                if abs_path not in self.file_id_mapping:
                    self.file_id_mapping[abs_path] = file_id
                    log.info(f"Added mapping for uploaded file: {abs_path} -> {file_id}")
                
                # Extraire et afficher le nom du service pour information
                service_name = self.extract_service_name(filename)
                log.info(f"File {filename} identified as service: {service_name}")
                
            # Analyser la structure
            response = self.analyze_slide_structure(self.chat_id)
            if "error" in response:
                response = f"Erreur lors de l'analyse de la structure: {response['error']}"
            else:
                # Formater la réponse
                response = self.format_all_slide_data(response)
                # Stocker la structure en cache
                self.cached_structure = response
                
            self.system_prompt = "# Voici les informations des fichiers PPTX toutes les informations sont importantes pour la compréhension du message de l'utilisateur et les données sont triées : \n\n" +  response + "# voici le message de l'utilisateur : " 
            
            # Save file mappings after processing new files
            self.save_file_mappings()
        
        return body

    def get_existing_summaries(self, folder_name=None):
        """
        Récupère la liste des fichiers de résumé existants pour le chat_id actuel et les télécharge vers OpenWebUI.
        Cherche les fichiers dans les dossiers OUTPUT et pptx_folder.
        
        Args:
            folder_name (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            list: Liste des tuples (filename, url) des résumés.
        """
        if folder_name is None:
            folder_name = self.chat_id
        log.info(f"Getting existing summaries for folder: {folder_name}")
        
        summaries = []
        
        # Check in the OUTPUT folder first (traditional location for summaries)
        output_path = os.path.join(OUTPUT_FOLDER, folder_name)
        log.info(f"Checking output path: {output_path}")
        
        # Ensure OUTPUT directory exists
        os.makedirs(output_path, exist_ok=True)
        
        # Also check in pptx_folder (new location for reports)
        pptx_path = os.path.join(UPLOAD_FOLDER, folder_name)
        log.info(f"Checking pptx folder path: {pptx_path}")
        
        # Ensure pptx directory exists
        os.makedirs(pptx_path, exist_ok=True)
        
        try:
            # Check for PPTX files in OUTPUT directory
            if os.path.exists(output_path):
                files = os.listdir(output_path)
                log.info(f"Files in output folder: {files}")
                
                for filename in files:
                    if filename.endswith(".pptx"):
                        log.info(f"Found summary file in OUTPUT: {filename}")
                        file_path = os.path.join(output_path, filename)
                        
                        # Upload file to OpenWebUI to get a download URL
                        upload_result = self.download_file_openwebui(file_path)
                        
                        if "error" in upload_result:
                            log.error(f"Error uploading file to OpenWebUI: {upload_result['error']}")
                            continue
                        
                        download_url = upload_result.get("download_url", "")
                        if download_url:
                            log.info(f"Generated OpenWebUI URL: {download_url}")
                            summaries.append(("OUTPUT/" + filename, download_url))
                        else:
                            log.error(f"No download URL received for file: {filename}")
            
            # Check for PPTX files in pptx_folder directory
            if os.path.exists(pptx_path):
                files = os.listdir(pptx_path)
                log.info(f"Files in pptx folder: {files}")
                
                for filename in files:
                    if filename.endswith(".pptx") and ("_summary_" in filename or "_text_summary_" in filename):
                        log.info(f"Found summary/report file in pptx_folder: {filename}")
                        file_path = os.path.join(pptx_path, filename)
                        
                        # Upload file to OpenWebUI to get a download URL
                        upload_result = self.download_file_openwebui(file_path)
                        
                        if "error" in upload_result:
                            log.error(f"Error uploading file to OpenWebUI: {upload_result['error']}")
                            continue
                        
                        download_url = upload_result.get("download_url", "")
                        if download_url:
                            log.info(f"Generated OpenWebUI URL: {download_url}")
                            summaries.append(("pptx_folder/" + filename, download_url))
                        else:
                            log.error(f"No download URL received for file: {filename}")
            
            # Save any new mappings we've created
            if len(summaries) > 0:
                self.save_file_mappings()
                
            log.info(f"Final summaries list: {summaries}")
        except Exception as e:
            log.error(f"Error listing summary files: {str(e)}")
            log.error(f"Current working directory: {os.getcwd()}")
        
        return summaries

    def pipe(self, body: dict, user_message: str, model_id: str, messages: List[dict]) -> Generator[str, None, None]:
        """
        Gère le pipeline de traitement des messages et des commandes spécifiques.

        Cette méthode traite différentes commandes comme /summarize, /structure, et /clear, 
        et gère le streaming de réponses du modèle.

        Args:
            body (dict): Le corps de la requête contenant des métadonnées.
            user_message (str): Le message de l'utilisateur.
            model_id (str): L'identifiant du modèle utilisé.
            messages (List[dict]): Liste des messages précédents.

        Yields:
            str: Réponses formatées en Server-Sent Events (SSE) compatibles avec OpenWebUI.

        Commandes supportées:
        - /summarize: Tente de résumer les fichiers PPTX
        - /structure: Analyse la structure des diapositives
        - /clear: Supprime tous les fichiers de la conversation
        """
        message = user_message.lower()  # Convertir en minuscules pour simplifier la correspondance
        __event_emitter__ = body.get("__event_emitter__")

        # Check if we're waiting for confirmation
        if self.waiting_for_confirmation:
            if message in ["yes", "y", "oui", "o"]:
                self.waiting_for_confirmation = False
                
                # If we were waiting for summarize confirmation
                if self.confirmation_command == "summarize":
                    # Generate a new summary
                    log.info(f"Generating summary with additional info: {self.confirmation_additional_info}")
                    response = self.summarize_folder(add_info=self.confirmation_additional_info)
                    if "error" in response:
                        response = f"Erreur lors de la génération du résumé: {response['error']}"
                    else:
                        response = f"Le résumé de tous les fichiers a été généré avec succès. URL de téléchargement: \n{response.get('download_url', 'Non disponible')}"
                    
                    # Save mappings after generating a new summary
                    self.save_file_mappings()
                    
                    yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
                    yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                    self.last_response = response
                    return
            
            elif message in ["no", "n", "non"]:
                self.waiting_for_confirmation = False
                response = "Génération de résumé annulée."
                yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = response
                return
            
            # Reset if we get any other input
            self.waiting_for_confirmation = False
        
        # Gestion des commandes spécifiques
        if "/summarize" in message:
            # Extract additional information after the /summarize command
            additional_info = None
            if " " in message:
                command_parts = message.split(" ", 1)
                if len(command_parts) > 1 and command_parts[1].strip():
                    additional_info = command_parts[1].strip()
            
            # Get existing summaries
            existing_summaries = self.get_existing_summaries()
            log.info(f"ACRA - Pipeline: Existing summaries: {existing_summaries}")
            
            if existing_summaries:
                response = "Voici les résumés existants pour cette conversation:\n\n"
                for filename, url in existing_summaries:
                    response += f"- {filename}: {url}\n"
                
                response += "\nVoulez-vous générer un nouveau résumé? (Oui/Non)"
                
                # Set state to wait for confirmation
                self.waiting_for_confirmation = True
                self.confirmation_command = "summarize"
                self.confirmation_additional_info = additional_info
                
                yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = response
                return
            else:
                # No existing summaries, generate one directly
                response = self.summarize_folder(add_info=additional_info)
                if "error" in response:
                    response = {"error": f"Erreur lors de la génération du résumé: {response['error']}"}
                else:
                    introduction_prompt = f"""Tu es un assistant qui va générer une introduction pour un enssemble de fichiers PPTX je veux juste une description globale des fichiers impliqués dans le message de 
                l'utilisateur pas de cas par cas et sourtout quelque chose de consit et renvoie uniquement l'introduction (pas d'explication) si tu vois une information importante ou une alerte critique, tu dois 
                la signaler dans l'introduction. Voici le contenu de tous les fichiers : {self.system_prompt} Tu dois renvoyer uniquement l'introduction (pas d'explication).
                """
                    introduction = self.small_model.invoke(introduction_prompt if "introduction_prompt" in locals() else self.system_prompt)
                    response = f"{introduction}\n\n Le résumé de tous les fichiers a été généré avec succès.\n\n  ### URL de téléchargement: \n{response.get('download_url', 'Non disponible')}"
                    
                    # Save mappings after generating a new summary
                    self.save_file_mappings()
                
                yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = response
                return
        
        elif "/structure" in message:
            if self.cached_structure is None:
                # Récupérer la structure des diapositives
                response = self.analyze_slide_structure(self.chat_id)
                
                if "error" in response:
                    response_text = f"Erreur lors de l'analyse de la structure: {response['error']}"
                    if __event_emitter__:
                        __event_emitter__({"type": "content", "content": response_text})
                    yield f"data: {json.dumps({'choices': [{'message': {'content': response_text}}]})}\n\n"
                    yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                    self.last_response = response_text
                    return
                
                # Formater les données de la structure
                formatted_response = self.format_all_slide_data(response)
                self.cached_structure = formatted_response
                if __event_emitter__:
                    __event_emitter__({"type": "content", "content": formatted_response})
                yield f"data: {json.dumps({'choices': [{'message': {'content': formatted_response}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = formatted_response
                return
            else:
                if __event_emitter__:
                    __event_emitter__({"type": "content", "content": self.cached_structure})
                yield f"data: {json.dumps({'choices': [{'message': {'content': self.cached_structure}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = self.cached_structure
                return
        elif "/generate" in message:
            # Extraire le texte après la commande
            text_content = user_message.replace("/generate", "").strip()
            if not text_content:
                response = "Veuillez fournir du texte après la commande /generate pour générer un rapport."
            else:
                # On utilise la méthode generate_report qui maintenant fait un POST avec le texte dans le body
                response = self.generate_report(self.chat_id, text_content)
                if "error" in response:
                    response = f"Erreur lors de la génération du rapport: {response['error']}"
                else:
                    response = f"Le rapport a été généré avec succès à partir du texte fourni.\n\n### URL de téléchargement:\n{response.get('download_url', 'Non disponible')}"
                    # Save mappings after generating a new report
                    self.save_file_mappings()
            
            if __event_emitter__:
                __event_emitter__({"type": "content", "content": response})
            yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.last_response = response
            return
        
        elif "/clear" in message:
            response = self.delete_all_files()
            if "error" in response:
                response = f"Erreur lors de la suppression des fichiers: {response['error']}"
            else:
                response = response.get('message', "Les fichiers ont été supprimés avec succès.")
                self.file_path_list = []  # Réinitialiser la liste des fichiers
                self.cached_structure = None  # Réinitialiser la structure en cache
            if __event_emitter__:
                __event_emitter__({"type": "content", "content": response})
            yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            self.last_response = response
            return

        elif "/merge" in message:
            folderpath = os.path.join("./pptx_folder", self.chat_id)
            response = merge_pptx_files(folderpath, os.path.join("./OUTPUT", self.chat_id, "merged_presentation.pptx"))
            if "error" in response:
                response = f"Erreur lors de la fusion des fichiers: {response['error']}"
            else:
                response = "Les fichiers ont été fusionnés avec succès." + response
                yield f"data: {json.dumps({'choices': [{'message': {'content': response}}]})}\n\n"
                yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
                self.last_response = response
                return
        # Ajouter la dernière réponse au contexte si elle existe
        if user_message:
            user_message += f"\n\n *Last response generated :* {self.last_response}"
        else:
            # Afficher les commandes disponibles si aucune réponse précédente
            commands = """Les commandes sont les suivantes : \n
            /summarize [instructions] --> Affiche les résumés existants et demande confirmation avant d'en générer un nouveau. Vous pouvez ajouter des instructions spécifiques après la commande pour guider le résumé.
            /structure --> Renvoie la structure des fichiers 
            /clear --> Retire tous les fichiers de la conversation
            /generate --> genere tout le pptx en fonction du texte ( /generate [Avancements de la semaine])
            /merge --> Fusionne tous les fichiers pptx envoyés
            """
            self.last_response = commands
            yield f"data: {json.dumps({'choices': [{'message': {'content': commands}}]})}\n\n"
            yield f"data: {json.dumps({'choices': [{'finish_reason': 'stop'}]})}\n\n"
            return
        
        # Initialiser le contenu cumulatif
        cumulative_content = ""
        user_message = self.system_prompt + "\n\n" + user_message
        
        try:
            # Format standard OpenAI-like qui est attendu par OpenWebUI
            # Premier message pour initialiser le stream
            yield f"data: {json.dumps({'choices': [{'delta': {'role': 'assistant'}}]})}\n\n"
            
            # Streamer la réponse depuis le modèle
            for chunk in self.streaming_model.stream(user_message):
                if isinstance(chunk, str):
                    content_delta = chunk
                else:
                    content_delta = chunk.content if hasattr(chunk, 'content') else str(chunk)
                
                # Nettoyer le contenu pour éviter les problèmes de formatage
                content_delta = content_delta.replace('\r', '')
                
                # Ajouter au contenu cumulatif
                cumulative_content += content_delta
                
                # Envoi de l'événement au client si un émetteur est disponible
                if __event_emitter__:
                    __event_emitter__({"type": "content_delta", "delta": content_delta})
                
                # Format compatible avec le standard OpenAI utilisé par OpenWebUI
                delta_response = {
                    "choices": [
                        {
                            "delta": {"content": content_delta}
                        }
                    ]
                }
                
                # Yield en format SSE (Server-Sent Events)
                yield f"data: {json.dumps(delta_response)}\n\n"
                
            # Message de fin spécifique
            yield f"data: {json.dumps({'choices': [{'delta': {}, 'finish_reason': 'stop'}]})}\n\n"
            yield f"data: [DONE]\n\n"  # Signal de fin standard OpenAI
            
        except Exception as e:
            error_message = f"Erreur lors du streaming de la réponse: {str(e)}"
            if __event_emitter__:
                __event_emitter__({"type": "error", "error": error_message})
            yield f"data: {json.dumps({'error': error_message})}\n\n"
            yield f"data: [DONE]\n\n"  # Même en cas d'erreur, on ferme proprement
            return
        
        self.last_response = cumulative_content

    def delete_file(self, file_path, update_mapping=True):
        """
        Supprime un fichier spécifique et met à jour le mapping si nécessaire.
        
        Args:
            file_path (str): Chemin du fichier à supprimer
            update_mapping (bool): Si True, met à jour le mapping des fichiers
            
        Returns:
            dict: Résultat de l'opération
        """
        try:
            abs_path = os.path.abspath(file_path)
            log.info(f"Deleting file: {abs_path}")
            
            if not os.path.exists(abs_path):
                log.warning(f"File not found: {abs_path}")
                return {"error": "Fichier non trouvé"}
            
            # Delete the file
            os.remove(abs_path)
            log.info(f"File deleted successfully: {abs_path}")
            
            # Update mapping if requested
            if update_mapping and abs_path in self.file_id_mapping:
                file_id = self.file_id_mapping[abs_path]
                del self.file_id_mapping[abs_path]
                log.info(f"Removed mapping for file {abs_path} with ID {file_id}")
                self.save_file_mappings()
            
            return {"message": f"Fichier {os.path.basename(file_path)} supprimé avec succès"}
        except Exception as e:
            log.error(f"Error deleting file {file_path}: {str(e)}")
            return {"error": f"Erreur lors de la suppression du fichier: {str(e)}"}

    def cleanup_orphaned_mappings(self):
        """
        Nettoie les mappings qui pointent vers des fichiers qui n'existent plus.
        
        Returns:
            int: Nombre de mappings supprimés
        """
        orphaned_mappings = []
        
        # Find all mappings pointing to files that no longer exist
        for file_path, file_id in list(self.file_id_mapping.items()):
            if not os.path.exists(file_path):
                orphaned_mappings.append((file_path, file_id))
        
        # Remove orphaned mappings
        removed_count = 0
        for file_path, file_id in orphaned_mappings:
            log.info(f"Removing orphaned mapping: {file_path} -> {file_id}")
            del self.file_id_mapping[file_path]
            removed_count += 1
        
        # Save updated mappings if any were removed
        if removed_count > 0:
            log.info(f"Removed {removed_count} orphaned mappings")
            self.save_file_mappings()
        else:
            log.info("No orphaned mappings found")
        
        return removed_count

pipeline = Pipeline()

if __name__ == "__main__":
    summarize_ppt("1040706a-776f-4233-b823-b49658dc42dd")