import json
import os
import sys
import shutil
import requests
from typing import List, Union, Generator, Iterator, Dict, Any
from langchain_ollama import  OllamaLLM
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..","src")))

from OLLibrary.utils.text_service import remove_tags_keep



class Pipeline:

    def __init__(self):

        self.last_response = None

        # self.model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", num_ctx=32000)
        self.streaming_model = OllamaLLM(model="deepseek-r1:14b", base_url="http://host.docker.internal:11434", num_ctx=64000, stream=True)

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

    async def inlet(self, body: dict, user: dict) -> dict:
        print(f"Received body: {body}")
        
        # Get conversation ID from body
        if body.get("metadata", {}).get("chat_id") != None:
            self.chat_id = body.get("metadata", {}).get("chat_id", "default")
            if not self.current_chat_id:
                self.current_chat_id = self.chat_id

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

        # Gestion des commandes spécifiques (/summarize, /structure, /clear)
        if "/summarize" in message:
            # No filename provided, summarize all files by default
            response = self.summarize_folder()
            if "error" in response:
                response = f"Erreur lors de la génération du résumé: {response['error']}"
            else:
                response = f"Le résumé de tous les fichiers a été généré avec succès. URL de téléchargement: {response.get('download_url', 'Non disponible')}"
            
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
