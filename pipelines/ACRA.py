import json
import os
import shutil
import requests
from typing import List, Union, Generator, Iterator
from langchain_ollama import  OllamaLLM
from pydantic import BaseModel



class Pipeline:

    class Valves(BaseModel): 
        LLAMAINDEX_OLLAMA_BASE_URL: str = "http://host.docker.internal:11434"
        LLAMAINDEX_MODEL_NAME: str = "gemma3:27b"

    def __init__(self):

        self.valves = self.Valves(
            **{
                "LLAMAINDEX_OLLAMA_BASE_URL": os.getenv("LLAMAINDEX_OLLAMA_BASE_URL", "http://host.docker.internal:11434"),
                "LLAMAINDEX_MODEL_NAME": os.getenv("LLAMAINDEX_MODEL_NAME", "gemma3:27b"),
            }
        )
        
        self.last_response = None

        self.model = OllamaLLM(model=self.valves.LLAMAINDEX_MODEL_NAME, base_url=self.valves.LLAMAINDEX_OLLAMA_BASE_URL)
        
        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030"

        self.file_path_list = []

        self.chat_id = ""
    

    def fetch(self, endpoint):
            """Effectue une requête GET synchrone"""
            url = f"{self.api_url}/{endpoint}"
            response = requests.get(url)
            return response.json() if response.status_code == 200 else {"error": "Request failed"}

    def summarize_presentation(self, filename):
        return self.fetch(f"acra/{filename}")

    def analyze_slide_structure(self, filename):
        return self.fetch(f"get_slide_structure/{filename}")
    
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
            return "❌ Aucun fichier PPTX fourni."

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

            for item in content:
                temp_global_content.append(f"{item} : {content.get(item).get('information')}")
                if content.get(item).get("alerts"):
                    alerts = content.get(item).get("alerts")
                    temp_alerts_critical.extend(f"**{item}**\n{alert}\n" for alert in alerts.get("critical_alerts", []))
                    temp_alerts_warning.extend(f"**{item}**\n{alert}\n" for alert in alerts.get("small_alerts", []))
                    temp_alerts_advancements.extend(f"**{item}**\n{alert}\n" for alert in alerts.get("advancements", [])) 

            result += f"Informations globales : {temp_global_content}\n"
            result += f"🔴 **Alertes Critiques :**\n{temp_alerts_critical}\n"
            result += f"🟡 **Alertes à surveiller :**\n{temp_alerts_warning}\n"
            result += f"🟢 **Avancements :**\n{temp_alerts_advancements}\n"
            
            result += "-" * 50 + "\n"  # Séparateur entre fichiers

        return result.strip()


    def delete_all_files(self):
        url = f"{self.api_url}/delete_all_pptx_files"
        response = requests.delete(url) 
        print(response)

        return response
    

    async def inlet(self, body: dict, user: dict) -> dict:
        print(f"Received body: {body}")
        
        # Extraction des informations de fichiers depuis body['metadata']['files']
        files = body.get("metadata", {}).get("files", [])
        if files:
            for file_entry in files:
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                file_id = file_data.get("id", "N/A")

                filecomplete_name = file_id + "_" + filename

                source_path = os.path.join("./uploads", filecomplete_name)
                # Chemin de destination dans le dossier pptx_folder
                destination_path = os.path.join("./pptx_folder", filecomplete_name)
                
                self.file_path_list.append(destination_path)
                shutil.copy(source_path, destination_path)
        
        return body


    def pipe(
            self, body: dict, user_message: str, model_id: str, messages: List[dict]
        ) -> Union[str, Generator, Iterator]:
    
        message = user_message.lower()  # Convert to lowercase for easier matching
        # last_response = self.last_response

        # Check for commands anywhere in the message
        if "/summarize" in message:
            response = "YESSSS JE SUMMARIZE"
            self.last_response = response
            return response
            
        elif "/structure" in message:
            print('structure')
            # print("chat id : ",self.chat_id)
            response = self.fetch(f"get_slide_structure/")
            response = self.format_all_slide_data(response)
            # self.last_response = response
            return response
            
        elif "/clear" in message:
            self.delete_all_files()
            response = "all the files are clear import new files for the ACRA to work properly :)"
            self.last_response = response
            return response
            
        # Only use Ollama for non-command messages
        # if last_response:
        #     message += f"\n\n *Last response generated :* {last_response}"
        # response = self.model.invoke(message)
        # self.last_response = response
        return response
    
pipeline = Pipeline()
