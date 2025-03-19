import os
import shutil
import requests
from typing import List, Union, Generator, Iterator
from langchain_ollama import  OllamaLLM
from pydantic import BaseModel



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

        self.model = OllamaLLM(model="gemma3:27b", base_url="http://host.docker.internal:11434")
        
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
            return "Aucun fichier PPTX fourni."

        for presentation in presentations["presentations"]:
            filename = presentation.get("filename", "Unknown File")
            total_slides = presentation.get("slide data", {}).get("total_slides", 0)
            slides = presentation.get("slide data", {}).get("slides", [])

            result += f"\n📂 **Présentation : {filename}**\n"
            result += f"📊 **Nombre total de diapositives : {total_slides}**\n\n"

            for slide in slides:
                slide_number = slide.get("slide_number", "N/A")
                result += f"📄 **Diapositive {slide_number} :**\n"

                for shape in slide.get("shapes", []):
                    shape_type = shape.get("type", "Shape")
                    
                    # Si c'est du texte classique
                    if shape_type == "Shape" and "text" in shape:
                        result += f"🔹 **Texte :** {shape['text']}\n"

                    # Si c'est un tableau
                    elif shape_type == "GraphicFrame" and "table" in shape:
                        result += "📊 **Tableau :**\n"
                        for row in shape["table"]:
                            row_text = " | ".join(row).strip()
                            if row_text:  # Évite d'afficher des lignes vides
                                result += f"   - {row_text}\n"
                
                result += "\n"  # Ajoute un espace entre les diapositives
            
            result += "-" * 50 + "\n"  # Séparateur entre fichiers

        return result.strip()


    def delete_all_files(self,folder):
        url = f"{self.api_url}/delete_all_pptx_files{folder}"
        response = requests.delete(url) 
        print(response)

        return response
    

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

        # Check for commands anywhere in the message
        if "/summarize" in message:
            response = "YESSSS JE SUMMARIZE"
            self.last_response = response
            return response
            
        elif "/structure" in message:
            print('structure')
            print("chat id : ",self.chat_id)
            request_url = f"get_slide_structure/{self.chat_id}"
            print("request : " , request_url )
            response = self.fetch(request_url)
            print("response : " , response)
            response = self.format_all_slide_data(response)
            self.last_response = response
            return response
            
        elif "/clear" in message:
            response = self.delete_all_files(self.chat_id).get('message')
            self.last_response = response
            return response
            
        # Only use Ollama for non-command messages
        if last_response:
            message += f"\n\n *Last response generated :* {last_response}"
        response = self.model.invoke(message)
        self.last_response = response
        return response
    
pipeline = Pipeline()


# async def inlet(self, body: dict, user: dict) -> dict:
#     """Modifies form data before the OpenAI API request."""

#     # Extract file info for all files in the body
#     # here i have created an inmemory dictionary to link users to their owned files
#     file_info = self._extract_file_info(body)
#     self.file_contents[user["id"]] = file_info
#     return body
# async def analyze_slide_structure(self, filename):
    #     response = await self._make_request(f"{self.api_url}/get_slide_structure/{filename}", "GET")
    #     return response

    # async def analyze_slide_structure_with_color(self, filename):
    #     response = await self._make_request(f"{self.api_url}/get_slide_structure_wcolor/{filename}", "GET")
    #     return response

    # async def summarize_presentation(self, file_path):
    #     # Vérifier que le fichier existe
    #     if not os.path.exists(file_path):
    #         return {"error": "Fichier introuvable"}
    #     with open(file_path, 'rb') as f:
    #         files = {'file': f}
    #         response = await self._make_request(f"{self.api_url}/acra/", "POST", files=files)
    #         return response

    # async def download_file(self, filename):
    #     response = await self._make_request(f"{self.api_url}/download/{filename}", "GET", stream=True)
    #     return response

    # async def _make_request(self, url, method, **kwargs):
    #     async with aiohttp.ClientSession() as session:
    #         if method == "GET":
    #             async with session.get(url, **kwargs) as response:
    #                 return await response.json()
    #         elif method == "POST":
    #             async with session.post(url, **kwargs) as response:
    #                 return await response.json()


# async def pipe(self, user_message: str, file_input: str = None):
    #     """
    #     Cette méthode détermine quelle action réaliser en fonction du message utilisateur
    #     et du fichier fourni.

    #     Args:
    #         user_message (str): Le message de l'utilisateur, devant contenir par exemple "analyze" ou "summarize".
    #         file_input (str, optional): Le chemin ou le nom du fichier à traiter.

    #     Returns:
    #         dict: La réponse obtenue via l'API.
    #     """
    #     # Vérifier qu'un fichier est bien fourni
    #     if not file_input:
    #         return {"error": "Aucun fichier fourni."}
    #     print("caca")
    #     message = user_message.lower()
    #     print(message)
    #     if "summarize" in message:
    #         # On lance la summarization si le message contient "summarize"
    #         return await self.summarize_presentation(file_input)
    #     elif "analyze" in message:
    #         # Si "analyze" est présent, on peut choisir l'analyse avec ou sans couleur
    #         if "color" in message or "couleur" in message:
    #             return await self.analyze_slide_structure_with_color(file_input)
    #         else:
    #             return await self.analyze_slide_structure(file_input)
    #     else:
    #         return {"error": "Commande non reconnue. Veuillez inclure 'analyze' ou 'summarize' dans votre message."}


# async def save_text_file(body: dict, save_folder: str):
#     """Extrait et enregistre un fichier texte depuis le body."""
#     os.makedirs(save_folder, exist_ok=True)  # Créer le dossier s'il n'existe pas

#     for file in body.get("files", []):
#         file_name = file["file"]["filename"]
#         file_content = file["file"]["data"]["content"]  # Texte brut

#         file_path = os.path.join(save_folder, file_name)
#         with open(file_path, "w", encoding="utf-8") as f:
#             f.write(file_content)

#         print(f"✅ Fichier enregistré : {file_path}")
#         return file_path

# async def save_binary_file(body: dict, save_folder: str):
#     """Extrait et enregistre un fichier binaire (ex: .pptx) depuis le body."""
#     os.makedirs(save_folder, exist_ok=True)  # Créer le dossier s'il n'existe pas

#     for file in body.get("files", []):
#         file_name = file["file"]["filename"]
#         file_content = file["file"]["data"]["content"]  # Encodé en base64 ?

#         try:
#             # Décoder le fichier s'il est en base64
#             file_bytes = base64.b64decode(file_content)
#         except Exception:
#             print("⚠️ Le fichier n'est pas encodé en base64. Enregistrement brut.")
#             file_bytes = file_content.encode("utf-8")

#         file_path = os.path.join(save_folder, file_name)
#         with open(file_path, "wb") as f:
#             f.write(file_bytes)

#         print(f"✅ Fichier enregistré : {file_path}")
#         return file_path