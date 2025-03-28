import json
import os
import sys
import shutil
import requests
from typing import List, Union, Generator, Iterator
from langchain_ollama import  OllamaLLM
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..","src")))

from OLLibrary.utils.text_service import remove_tags_keep



class Pipeline:

    # class Valves(BaseModel): 
    #     LLAMAINDEX_OLLAMA_BASE_URL: str = "http://host.docker.internal:11434"
    #     LLAMAINDEX_MODEL_NAME: str = "gemma3:27b"

    def __init__(self):

        # self.valves = self.Valves(
        #     **{
                
        #     }
        # )
        
        self.last_response = None

        self.model = OllamaLLM(model="deepseek-r1:8b", base_url="http://host.docker.internal:11434", num_ctx=32000)
        
        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030"

        self.file_path_list = []

        self.chat_id = ""
    

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
            result += f"📊 **Nombre total de diapositives : {total_slides}**\n\n"

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
        
        return body


    def pipe(
            self, body: dict, user_message: str, model_id: str, messages: List[dict]
        ) -> Union[str, Generator, Iterator]:
    
        message = user_message.lower()  # Convert to lowercase for easier matching
        last_response = self.last_response

        # Validate that we have a chat_id
        if not self.chat_id:
            response = "Erreur: Impossible d'identifier la conversation. Veuillez réessayer."
            self.last_response = response
            return response

        # Check for commands anywhere in the message
        if "/summarize" in message:
            # No filename provided, summarize all files by default
            response = self.summarize_folder()
            if "error" in response:
                response = f"Erreur lors de la génération du résumé: {response['error']}"
            else:
                response = f"Le résumé de tous les fichiers a été généré avec succès. URL de téléchargement: {response.get('download_url', 'Non disponible')}"
            
            self.last_response = response
            return response
            
        elif "/structure" in message:
            print('structure')
            print("chat id : ", self.chat_id)
            response = self.analyze_slide_structure()
            print("response : ", response)
            if "error" in response:
                response = f"Erreur lors de l'analyse de la structure: {response['error']}"
            else:
                response = self.format_all_slide_data(response)
            self.last_response = response
            return response
            
        elif "/clear" in message:
            response = self.delete_all_files()
            if "error" in response:
                response = f"Erreur lors de la suppression des fichiers: {response['error']}"
            else:
                response = response.get('message', "Les fichiers ont été supprimés avec succès.")
                self.file_path_list = []  # Clear the file path list
            self.last_response = response
            return response
            
        # Only use Ollama for non-command messages
        # Nettoyage et ajout du flag de raisonnement
        if model_id.startswith("reasoning/"):
            model_id = model_id.replace("reasoning/", "", 1)
        
        # Concaténer le dernier contexte (si disponible)
        if self.last_response:
            user_message += f"\n\n*Last response generated:* {self.last_response}"
        
        # Préparer le payload en incluant le flag de raisonnement
        payload = {
            "model": model_id,
            "prompt": user_message,
            "include_reasoning": True,
            # éventuellement d'autres paramètres spécifiques
        }
        
        # Exemple d'appel via une API ou méthode custom (ici à adapter selon OllamaLLM)
        response = self.model.invoke(user_message)  # ou un appel à l'API avec le payload préparé
        
        # Post-traitement pour ajouter le raisonnement dans la réponse finale
        if isinstance(response, dict) and "choices" in response:
            for choice in response["choices"]:
                if "message" in choice and "reasoning" in choice["message"]:
                    reasoning = choice["message"]["reasoning"] + "\n"
                    choice["message"]["content"] = f"{reasoning}{choice['message']['content']}"
            final_response = "\n".join([choice["message"]["content"] for choice in response["choices"]])
        else:
            final_response = response  # si c'est simplement une chaîne
        
        self.last_response = final_response
        return final_response
    
pipeline = Pipeline()
