import json
import os
import sys
import shutil
import requests
import uuid
from typing import List, Union, Generator, Iterator, Dict, Any
from langchain_ollama import  OllamaLLM
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..","src")))
from core import summarize_ppt, get_slide_structure, delete_all_pptx_files, generate_pptx_from_text
from services import merge_pptx

from OLLibrary.utils.text_service import remove_tags_keep
from OLLibrary.utils.log_service import setup_logging, get_logger

import logging

# Set up the main application logger
setup_logging(app_name="ACRA_Pipeline")
load_dotenv()
# Use a specific logger for this module
log = get_logger(__name__)

class Pipeline:
    def __init__(self):
        log.info("Initializing ACRA Pipeline")
        self.last_response = None

        self.use_api = False

        # self.model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", num_ctx=32000)
        self.streaming_model = OllamaLLM(model="deepseek-r1:14b", base_url="http://host.docker.internal:11434", num_ctx=131000, stream=True)

        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030"

        self.small_model = OllamaLLM(model="qwen2.5:3b", base_url="http://host.docker.internal:11434", num_ctx=131000)

        self.file_path_list = []

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
        log.info("ACRA Pipeline initialized successfully")

    def generate_report(self, foldername, info):
        """
        Génère un rapport à partir du texte fourni en utilisant une requête POST.
        
        Args:
            foldername (str): Le nom du dossier où stocker le rapport
            info (str): Le texte à analyser pour générer le rapport
            
        Returns:
            dict: Résultat de la requête avec l'URL de téléchargement
        """

        return generate_pptx_from_text(foldername, info)

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

    def summarize_folder(self, foldername=None):
        """
        Envoie une demande pour résumer tous les fichiers PowerPoint dans un dossier.
        
        Args:
            foldername (str, optional): Le nom du dossier à résumer. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'opération de résumé.
        """
        if foldername is None:
            foldername = self.chat_id
        
        if self.use_api:
            return self.fetch(f"acra/{foldername}")
        return summarize_ppt(foldername)

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
        
        if self.use_api:
            return self.fetch(f"get_slide_structure/{foldername}")
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
        Supprime tous les fichiers dans un dossier.
        
        Args:
            foldername (str, optional): Le nom du dossier à vider. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'opération de suppression.
        """
        if foldername is None:
            foldername = self.chat_id

        if self.use_api:
            url = f"{self.api_url}/delete_all_pptx_files/{foldername}"
            response = requests.delete(url) 
            return response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}
        return delete_all_pptx_files(foldername)
    
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

    async def inlet(self, body: dict, user: dict) -> dict:
        log.info(f"Received body: {body}")
        
        # Debug log the current state
        log.info(f"Current state - self.chat_id: '{self.chat_id}', self.current_chat_id: '{self.current_chat_id}'")
        
        # Get conversation ID from body
        if body.get("metadata", {}).get("chat_id") != None:
            self.chat_id = body.get("metadata", {}).get("chat_id", "default")
            if not self.current_chat_id:
                self.current_chat_id = self.chat_id

        # Create foldername with conversation ID
        conversation_folder = os.path.join("./pptx_folder", self.chat_id)
        os.makedirs(conversation_folder, exist_ok=True)
        print(f"Created folder at : {os.path.join('./pptx_folder', self.chat_id)}")

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
                
                # Extraire et afficher le nom du service pour information
                service_name = self.extract_service_name(filename)
                print(f"Fichier {filename} extrait comme service: {service_name}")
                
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
        
        return body

    def get_existing_summaries(self, folder_name=None):
        """
        Récupère la liste des fichiers de résumé existants pour le chat_id actuel.
        
        Args:
            folder_name (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            list: Liste des tuples (filename, url) des résumés.
        """
        if folder_name is None:
            folder_name = self.chat_id
        log.info(f"ACRA - Pipeline: Getting existing summaries for folder: {folder_name}")
        output_folder = os.getenv("OUTPUT_FOLDER", "OUTPUT")
        log.info(f"ACRA - Pipeline: Output folder: {output_folder}")
        summaries = []
        folder_path = os.path.join(output_folder, folder_name)
        log.info(f"ACRA - Pipeline: Folder path: {folder_path}")
        log.info(f"ACRA - Pipeline: Folder exists: {os.path.exists(folder_path)}")
        os.makedirs(folder_path, exist_ok=True)
        log.info(f"ACRA - Pipeline: Makedirs: {folder_path}")
        
        try:
            # List all files in the current chat folder
            files = os.listdir(folder_path)
            log.info(f"ACRA - Pipeline: All files in directory: {files}")
            for filename in files:
                log.info(f"ACRA - Pipeline: Processing file: {filename}")
                if filename and filename.endswith(".pptx"):
                    download_url = f"http://localhost:5050/download/{folder_name}/{filename}"
                    log.info(f"ACRA - Pipeline: Download URL: {download_url}")
                    summaries.append((filename, download_url))
            log.info(f"ACRA - Pipeline: Final summaries list: {summaries}")
        except Exception as e:
            log.error(f"ACRA - Pipeline: Error listing files: {str(e)}")
            log.error(f"ACRA - Pipeline: Current working directory: {os.getcwd()}")
            log.error(f"ACRA - Pipeline: Absolute folder path: {os.path.abspath(folder_path)}")
        
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
                    response = self.summarize_folder(additional_info=self.confirmation_additional_info)
                    if "error" in response:
                        response = f"Erreur lors de la génération du résumé: {response['error']}"
                    else:
                        response = f"Le résumé de tous les fichiers a été généré avec succès. URL de téléchargement: \n{response.get('download_url', 'Non disponible')}"
                    
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
                response = self.summarize_folder(additional_info=additional_info)
                if "error" in response:
                    response = f"Erreur lors de la génération du résumé: {response['error']}"
                else:
                    introduction_prompt = f"""Tu es un assistant qui va générer une introduction pour un enssemble de fichiers PPTX je veux juste une description globale des fichiers impliqués dans le message de 
                l'utilisateur pas de cas par cas et sourtout quelque chose de consit et renvoie uniquement l'introduction (pas d'explication) si tu vois une information importante ou une alerte critique, tu dois 
                la signaler dans l'introduction. Voici le contenu de tous les fichiers : {self.system_prompt} Tu dois renvoyer uniquement l'introduction (pas d'explication).
                """
                introduction = self.small_model.invoke(introduction_prompt)
                response = f"{introduction}\n\n Le résumé de tous les fichiers a été généré avec succès.\n\n  ### URL de téléchargement: \n{response.get('download_url', 'Non disponible')}"
                
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
            output_merge = "./OUTPUT/"+self.chat_id + "/merged/" 
            input_merge = "./pptx_folder/" + self.chat_id
            response = str(merge_pptx(input_merge,output_merge))
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

pipeline = Pipeline()

if __name__ == "__main__":
    summarize_ppt("1040706a-776f-4233-b823-b49658dc42dd")