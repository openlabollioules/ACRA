import os
import shutil
import requests
from typing import List, Union, Generator, Iterator
from langchain_ollama import  OllamaLLM
from pydantic import BaseModel



class Pipeline:

    class Valves(BaseModel): 
        LLAMAINDEX_OLLAMA_BASE_URL: str = "http://host.docker.internal:11434"
        LLAMAINDEX_MODEL_NAME: str = "gemma3:12b"

    def __init__(self):

        self.valves = self.Valves(
            **{
                "LLAMAINDEX_OLLAMA_BASE_URL": os.getenv("LLAMAINDEX_OLLAMA_BASE_URL", "http://host.docker.internal:11434"),
                "LLAMAINDEX_MODEL_NAME": os.getenv("LLAMAINDEX_MODEL_NAME", "gemma3:12b"),
            }
        )
        
        self.model = OllamaLLM(model=self.valves.LLAMAINDEX_MODEL_NAME, base_url=self.valves.LLAMAINDEX_OLLAMA_BASE_URL)
        
        self.api_url = "http://host.docker.internal:5050"

        self.openwebui_api = "http://host.docker.internal:3030"

        self.file_path_list = []
    

    def fetch(self, endpoint):
            """Effectue une requête GET synchrone"""
            url = f"{self.api_url}/{endpoint}"
            response = requests.get(url)
            return response.json() if response.status_code == 200 else {"error": "Request failed"}

    def summarize_presentation(self, filename):
        return self.fetch(f"acra/{filename}")

    def analyze_slide_structure(self, filename):
        return self.fetch(f"get_slide_structure/{filename}")
    
    def format_slide_data(self, slide_data: dict) -> str:
        filename = slide_data.get("filename", "Unknown File")
        total_slides = slide_data.get("slide data", {}).get("total_slides", 0)
        slides = slide_data.get("slide data", {}).get("slides", [])

        result = f"📂 **Présentation : {filename}**\n"
        result += f"📊 **Nombre total de diapositives : {total_slides}**\n\n"

        for slide in slides:
            slide_number = slide.get("slide_number", "N/A")
            result += f"📄 **Diapositive {slide_number} :**\n"

            # Formater les shapes et graphiques
            for shape in slide.get("shapes", []):
                shape_type = shape.get("type", "Shape")
                if shape_type == "Shape" and "text" in shape:
                    result += f"🔹 **Texte :** {shape['text']}\n"

                elif shape_type == "GraphicFrame" and "table" in shape:
                    result += "📊 **Tableau :**\n"
                    for row in shape["table"]:
                        result += f"   - {' | '.join(row)}\n"
            
            result += "\n"

        return result.strip()


    

    async def on_startup(self):
        pass

    async def on_shutdown(self):
        pass
    

    # async def inlet(self, body: dict, user: dict) -> dict:
    #     """Modifies form data before the OpenAI API request."""

    #     # Extract file info for all files in the body
    #     # here i have created an inmemory dictionary to link users to their owned files
    #     file_info = self._extract_file_info(body)
    #     self.file_contents[user["id"]] = file_info
    #     return body
    
    async def inlet(self, body: dict, user: dict) -> dict:
        print(f"Received body: {body}")
        
        # Extraction des informations de fichiers depuis body['metadata']['files']
        files = body.get("metadata", {}).get("files", [])
        if files : 
            for file_entry in files:
                # Récupération des infos du fichier dans le dictionnaire "file"
                file_data = file_entry.get("file", {})
                filename = file_data.get("filename", "N/A")
                file_id = file_data.get("id", "N/A")

                print(f"Filename: {filename}")
                print(f"File ID: {file_id}")

                # Correction de la concaténation pour obtenir le nom complet du fichier
                filecomplete_name = file_id + "_" + filename

                # Chemin source du fichier dans le dossier uploads
                source_path = os.path.join("./uploads", filecomplete_name)
                # Chemin de destination dans le dossier pptx_folder
                destination_path = os.path.join("./pptx_folder", filecomplete_name)
                
                self.file_path_list.append(destination_path)
                # Copie du fichier
                shutil.copy(source_path, destination_path)
        
        return body


    def pipe(
            self, body: dict, user_message: str, model_id: str, messages: List[dict]
        ) -> Union[str, Generator, Iterator]:
    
        message = user_message + " respond in English only"

        # for files in self.file_path_list:
        # structure = self.fetch(f"get_slide_structure/CRA_SERVICE_CYBER.pptx")
        
        parts = user_message.strip().split(" ", 1)  # ["commande", "argument"]
        command = parts[0].lower()
        # argument = parts[1] if len(parts) > 1 else None

        if command == "/summarize" :
            # response = self.summarize_presentation()
            response ="YESSSS JE SUMMARIZE "
            print('summarize')
        elif command == "/structure":
            print('structure')
            response = self.fetch(f"get_slide_structure/CRA_SERVICE_CYBER.pptx")
            
            response = self.format_slide_data(response)
        else:
            response = {"error Commande invalide. Essayez /summarize <filename> ou /structure <filename>"}

        # print("Processed message:", message)

        # response = self.model.invoke(message)
        # response = response +f"\n {structure}"

        return response
    
pipeline = Pipeline()


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