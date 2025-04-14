import json
import os
import sys
import shutil
import requests
from typing import List, Union, Generator, Iterator, Dict, Any
from langchain_ollama import  OllamaLLM
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..","src")))

from OLLibrary.utils.text_service import remove_tags_keep
from OLLibrary.utils.log_service import setup_logging, get_logger
from services import cleanup_orphaned_folders
import logging
from dotenv import load_dotenv
# Set up the main application logger
setup_logging(app_name="ACRA_Pipeline")

# Use a specific logger for this module
log = get_logger(__name__)
load_dotenv()

OUTPUT = os.getenv('OUTPUT_FOLDER')
class Pipeline:

    def __init__(self):

        self.last_response = None

        # self.model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", num_ctx=32000)
        self.streaming_model = OllamaLLM(model="deepseek-r1:32b", base_url="http://host.docker.internal:11434", num_ctx=131000, stream=True)

        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030"

        self.file_path_list = []

        self.chat_id = ""
        self.current_chat_id = ""  # To track conversation changes

        self.system_prompt = ""
        self.message_id = 0

    def reset_conversation_state(self):
        """Réinitialise les états spécifiques à une conversation"""
        self.last_response = None
        self.system_prompt = ""
        self.file_path_list = []
        self.message_id = 0

    def fetch(self, endpoint):
            """Effectue une requête GET synchrone"""
            url = f"{self.api_url}/{endpoint}"
            response = requests.get(url)
            return response.json() if response.status_code == 200 else {"error": "Request failed"}

    def post(self, endpoint, data=None, files=None):
        """Effectue une requête POST synchrone"""
        url = f"{self.api_url}/{endpoint}"
        response = requests.post(url, data=data, files=files)
        return response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}

    def summarize_folder(self, folder_name=None):
        """
        Envoie une demande pour résumer tous les fichiers PowerPoint dans un dossier.
        
        Args:
            folder_name (str, optional): Le nom du dossier à résumer. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'opération de résumé.
        """
        if folder_name is None:
            folder_name = self.chat_id
            
        return self.fetch(f"acra/{folder_name}")

    def analyze_slide_structure(self, folder_name=None):
        """
        Analyse la structure des diapositives dans un dossier.
        
        Args:
            folder_name (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'analyse.
        """
        if folder_name is None:
            folder_name = self.chat_id
            
        return self.fetch(f"get_slide_structure/{folder_name}")
    
    def format_all_slide_data(self, presentations: dict) -> str:
        """
        Formate les données de plusieurs présentations PPTX en une seule chaîne de texte structurée.

        Args:
            presentations (dict): Dictionnaire contenant plusieurs fichiers et leurs données.

        Returns:
            str: Une chaîne de texte structurée listant les informations de chaque fichier PPTX.
        """
        result = ""

        if not presentations.get("presentations"):
            return "Aucun fichier PPTX fourni."

        for presentation in presentations["presentations"]:
            filename = presentation.get("filename", "Unknown File")
            total_slides = presentation.get("slide data", {}).get("total_slides", 0)

            result += f"\n📂 **Présentation : {filename}**\n"
            result += f"📊 **Nombre total de diapositives : {total_slides+1}**\n\n"

            temp_alerts_critical = []
            temp_alerts_warning = []
            temp_alerts_advancements = []
            temp_global_content = []
            content = presentation.get("project_data", {})
            activites = content.get('activities', {})
            evenements = content.get('upcoming_events', [])

            for item in activites:
                # Format global content with project name as a heading
                temp_global_content.append(f"**{item}**:\n{activites.get(item).get('information')}")
                
                if activites.get(item).get("alerts"):
                    alerts = activites.get(item).get("alerts")
                    
                    # Format critical alerts
                    if alerts.get("critical_alerts"):
                        temp_alerts_critical.append(f"**{item}**:")
                        for alert in alerts.get("critical_alerts", []):
                            temp_alerts_critical.append(f"- {alert}")
                    
                    # Format small alerts
                    if alerts.get("small_alerts"):
                        temp_alerts_warning.append(f"**{item}**:")
                        for alert in alerts.get("small_alerts", []):
                            temp_alerts_warning.append(f"- {alert}")
                    
                    # Format advancements
                    if alerts.get("advancements"):
                        temp_alerts_advancements.append(f"**{item}**:")
                        for advancement in alerts.get("advancements", []):
                            temp_alerts_advancements.append(f"- {advancement}")

            # Format global information section
            result += "**Informations globales:**\n"
            for info in temp_global_content:
                result += f"{info}\n\n"
            
            # Format alerts sections with better styling
            if temp_alerts_critical:
                result += "🔴 **Alertes Critiques:**\n"
                result += "\n".join(temp_alerts_critical) + "\n\n"
            else:
                result += "🔴 **Alertes Critiques:** Aucune alerte critique à signaler.\n\n"
                
            if temp_alerts_warning:
                result += "🟡 **Alertes à surveiller:**\n"
                result += "\n".join(temp_alerts_warning) + "\n\n"
            else:
                result += "🟡 **Alertes à surveiller:** Aucune alerte mineure à signaler.\n\n"
                
            if temp_alerts_advancements:
                result += "🟢 **Avancements:**\n"
                result += "\n".join(temp_alerts_advancements) + "\n\n"
            else:
                result += "🟢 **Avancements:** Aucun avancement significatif à signaler.\n\n"

            # Format upcoming events section
            result += "**Evénements des semaines à venir:**\n"
            if evenements:
                result += f"{evenements}\n\n"
            else:
                result += "Aucun événement particulier prévu pour les semaines à venir.\n\n"
            
            result += "-" * 50 + "\n"  # Séparateur entre fichiers

        return result.strip()


    def delete_all_files(self, folder=None):
        """
        Supprime tous les fichiers dans un dossier.
        
        Args:
            folder (str, optional): Le nom du dossier à vider. Si None, utilise le chat_id.
        
        Returns:
            dict: Les résultats de l'opération de suppression.
        """
        if folder is None:
            folder = self.chat_id
            
        url = f"{self.api_url}/delete_all_pptx_files/{folder}"
        response = requests.delete(url) 
        return response.json() if response.status_code == 200 else {"error": f"Request failed with status {response.status_code}: {response.text}"}
    
    def get_files_in_folder(self, folder_name=None):
        """
        Récupère la liste des fichiers dans un dossier.
        
        Args:
            folder_name (str, optional): Le nom du dossier à analyser. Si None, utilise le chat_id.
        
        Returns:
            list: Liste des noms de fichiers PPTX dans le dossier.
        """
        if folder_name is None:
            folder_name = self.chat_id
            
        folder_path = os.path.join("./pptx_folder", folder_name)
        if not os.path.exists(folder_path):
            return []
            
        return [f for f in os.listdir(folder_path) if f.lower().endswith(".pptx")]

    def notify_chat_id_change(self, chat_id):
        """
        Notifie le frontend du changement de chat_id via un message SSE.
        Ce message sera intercepté par le chatBot.js pour mettre à jour le PptxViewer.
        
        Args:
            chat_id (str): L'identifiant de conversation à transmettre.
        """
        try:
            log.info(f"Notifying frontend of chat ID change: {chat_id}")
            # Ce message sera envoyé au parent de l'iframe via window.parent.postMessage
            # Format spécifique pour notre système de notification
            return {
                "type": "sse",
                "data": json.dumps({
                    "type": "chat_id_changed",
                    "chatId": chat_id
                })
            }
        except Exception as e:
            log.error(f"Error notifying chat ID change: {str(e)}")
            return None

    async def inlet(self, body: dict, user: dict) -> dict:
        print(f"Received body: {body}")
        
        # Get conversation ID from body
        new_chat_id = body.get("metadata", {}).get("chat_id", "default")
        log.info(f"Extracted new_chat_id from metadata: '{new_chat_id}'")
        
        # Always compare to current_chat_id, as that's our tracking variable
        if new_chat_id != self.current_chat_id:
            log.info(f"CHAT ID CHANGE DETECTED: '{self.current_chat_id}' → '{new_chat_id}'")
            
            # Préparer la notification pour le frontend
            body["__notifications__"] = self.notify_chat_id_change(new_chat_id)
            
            # Skip initial setup (first conversation)
            if self.current_chat_id:  # Only run if we had a previous conversation
                log.info(f"Previous conversation existed, running cleanup...")
                
                # Run cleanup for orphaned folders
                try:
                    log.info("Running cleanup for orphaned folders...")
                    cleanup_result = cleanup_orphaned_folders()
                    log.info(f"Cleanup completed: {cleanup_result}")
                except Exception as e:
                    log.error(f"Error running cleanup: {str(e)}", exc_info=True)
                
                # Reset conversation state for the new conversation
                self.reset_conversation_state()
            else:
                log.info("First conversation, skipping cleanup")
            
            # Update current chat ID tracking
            self.current_chat_id = new_chat_id
        else:
            log.info(f"No chat_id change detected (still '{self.current_chat_id}')")
        
        # Always update self.chat_id for use in the rest of the pipeline
        self.chat_id = new_chat_id
        
        # Create folder with conversation ID
        conversation_folder = os.path.join("./pptx_folder", self.chat_id)
        os.makedirs(conversation_folder, exist_ok=True)

        # Extract files from body['metadata']['files']
        files = body.get("metadata", {}).get("files", [])
        if files:
            for file_entry in files:
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                file_id = file_data.get("id", "N/A")

                filecomplete_name = file_id + "_" + filename

                source_path = os.path.join("./uploads", filecomplete_name)
                # Update destination to use conversation folder
                destination_path = os.path.join(conversation_folder, filecomplete_name)
                
                self.file_path_list.append(destination_path)
                shutil.copy(source_path, destination_path)
            response = self.analyze_slide_structure()
            if "error" in response:
                response = f"Erreur lors de l'analyse de la structure: {response['error']}"
            else:
                response = self.format_all_slide_data(response)
            self.system_prompt = "# Voici les informations des fichiers PPTX toutes les informations sont importantes pour la compréhension du message de l'utilisateur et les données sont triée : \n\n" +  response + "# voici le message de l'utilisateur : " 
        
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
                    download_url = f"http://localhost:5050/download/{OUTPUT}/{folder_name}/{filename}"
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
        # # Vérifier si c'est une nouvelle conversation en examinant les métadonnées du corps
        # new_chat_id = body.get("metadata", {}).get("chat_id", "default")
        # if self.current_chat_id and new_chat_id != self.current_chat_id:
        #     print(f"New conversation detected in pipe: {new_chat_id} (was: {self.current_chat_id})")
        #     self.reset_conversation_state()
        #     self.current_chat_id = new_chat_id
        #     self.chat_id = new_chat_id
        # else:
        #     self.current_chat_id = new_chat_id
        #     self.chat_id = new_chat_id
        message = user_message.lower()  # Convertir en minuscules pour simplifier la correspondance
        __event_emitter__ = body.get("__event_emitter__")

        # Toujours envoyer une notification de chat_id au début de chaque message
        # Cela garantit que le frontend est synchronisé avec le chat_id actuel
        if __event_emitter__ and self.chat_id:
            notification = self.notify_chat_id_change(self.chat_id)
            if notification and notification.get("type") == "sse":
                notification_data = notification.get("data", "{}")
                sse_event = f"data: {notification_data}\n\n"
                log.info(f"Sending initial chat_id notification: {sse_event}")
                yield sse_event

        # Check if there are notifications to send to the frontend
        if __event_emitter__ and body.get("__notifications__"):
            notification = body.get("__notifications__")
            log.info(f"Emitting notification: {notification}")
            
            # Si c'est une notification SSE, format spécial
            if notification.get("type") == "sse":
                notification_data = notification.get("data", "{}")
                # Format de l'événement SSE
                sse_event = f"data: {notification_data}\n\n"
                log.info(f"Sending SSE notification: {sse_event}")
                yield sse_event
            else:
                # Sinon, notification normale via event_emitter
                __event_emitter__(notification)

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
            response = self.analyze_slide_structure()
            if "error" in response:
                response = f"Erreur lors de l'analyse de la structure: {response['error']}"
            else:
                response = self.format_all_slide_data(response)
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
            if __event_emitter__:
                __event_emitter__({"type": "content", "content": response})
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
            /summarize --> Résume tous les fichiers pptx envoyé  
            /structure --> Renvoie la structure des fichiers 
            /clear --> Retire tous les fichiers de la conversation
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
