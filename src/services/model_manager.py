"""
Model Management Service for ACRA
Centralized LLM interactions and model configurations
"""
from typing import Generator, Any
from langchain_ollama import OllamaLLM
from OLLibrary.utils.log_service import get_logger
from OLLibrary.utils.text_service import remove_tags_no_keep
from config_pipeline import acra_config

log = get_logger(__name__)

class ModelManager:
    """
    Centralized model management for ACRA pipeline.
    Handles LLM initialization, configuration, and interactions.
    """
    
    def __init__(self):
        self._streaming_model = None
        self._small_model = None
        self._initialize_models()
    
    def _initialize_models(self):
        """Initialize the LLM models with configuration"""
        try:
            base_url = acra_config.get("OLLAMA_BASE_URL")
            
            # Initialize streaming model
            self._streaming_model = OllamaLLM(
                model=acra_config.get("STREAMING_MODEL"),
                base_url=base_url,
                num_ctx=acra_config.get("MODEL_CONTEXT_SIZE"),
                stream=True
            )
            
            # Initialize small model
            self._small_model = OllamaLLM(
                model=acra_config.get("SMALL_MODEL"),
                base_url=base_url,
                num_ctx=acra_config.get("SMALL_MODEL_CONTEXT_SIZE"),
                stream=True
            )
            
            log.info("Models initialized successfully")
            log.info(f"Streaming model: {acra_config.get('STREAMING_MODEL')}")
            log.info(f"Small model: {acra_config.get('SMALL_MODEL')}")
            
        except Exception as e:
            log.error(f"Error initializing models: {str(e)}")
            raise
    
    @property
    def streaming_model(self) -> OllamaLLM:
        """Get the streaming model instance"""
        if self._streaming_model is None:
            self._initialize_models()
        return self._streaming_model
    
    @property
    def small_model(self) -> OllamaLLM:
        """Get the small model instance"""
        if self._small_model is None:
            self._initialize_models()
        return self._small_model
    
    def invoke_small_model(self, prompt: str, clean_thinking: bool = True) -> str:
        """
        Invoke the small model with a prompt and optionally clean thinking tags.
        
        Args:
            prompt (str): The prompt to send to the model
            clean_thinking (bool): Whether to remove <think></think> tags from response
            
        Returns:
            str: The model's response
        """
        try:
            response = self.small_model.invoke(prompt)
            
            if clean_thinking:
                response = remove_tags_no_keep(response, '<think>', '</think>')
            
            return response.strip()
        except Exception as e:
            log.error(f"Error invoking small model: {str(e)}")
            raise
    
    def stream_response(self, prompt: str) -> Generator[str, None, None]:
        """
        Stream a response from the streaming model.
        
        Args:
            prompt (str): The prompt to send to the model
            
        Yields:
            str: Chunks of the model's response
        """
        try:
            for chunk in self.streaming_model.stream(prompt):
                if isinstance(chunk, str):
                    content_delta = chunk
                else:
                    content_delta = chunk.content if hasattr(chunk, 'content') else str(chunk)
                
                # Clean content for formatting issues
                content_delta = content_delta.replace('\r', '')
                yield content_delta
                
        except Exception as e:
            log.error(f"Error streaming response: {str(e)}")
            yield f"Error: {str(e)}"
    
    def extract_service_name(self, filename: str) -> str:
        """
        Extract service name from PowerPoint filename using the small model.
        
        Args:
            filename (str): The PowerPoint filename
            
        Returns:
            str: The extracted service name
        """
        prompt = f"""Tu es un assistant spécialisé dans le traitement automatique des noms de fichiers. 
        On te donne un nom de fichier de présentation (PowerPoint) contenant un identifiant unique suivi du titre du document. 
        Ton objectif est d'extraire uniquement le titre du document dans un format propre et lisible pour un humain. 
        Le titre est toujours situé après le dernier underscore (`_`) ou après une chaîne d'identifiants. 
        Supprime l'extension `.pptx` ou toute autre extension. 
        Remplace les underscores (`_`) ou tirets (`-`) par des espaces, et capitalise correctement chaque mot. 
        
        Exemple : 
        **Nom de fichier :** `dc56be63-37a6-4ed6-9223-50f545028ab4_CRA_SERVICE_UX.pptx`   
        **Titre extrait :** `Service UX` 
        
        Donne uniquement le titre extrait (pas d'explication), en une seule ligne. 
        Voici le nom du fichier : {filename}"""
        
        return self.invoke_small_model(prompt)
    
    def generate_project_grouping(self, project_names: list) -> list:
        """
        Generate project grouping suggestions using the small model.
        
        Args:
            project_names (list): List of project names to analyze
            
        Returns:
            list: List of project groups to merge
        """
        prompt = f"""Tu es un assistant qui analyse des noms de projets pour identifier ceux qui sont similaires ou liés.
        
        Voici les noms de projets :
        {project_names}
        
        Ton travail est d'identifier les groupes de projets qui ont des noms similaires ou qui font partie du même projet principal.
        
        RÈGLES:
        1. Regroupe les projets qui ont des noms similaires (ex: "NU" et "NU Gate" font partie du même groupe)
        2. Regroupe les projets qui ont des parties communes (ex: "NU" et "NU_MAX" font partie du même groupe)
        3. Deux projets qui ont le même noms mais avec des majuscules ou non ou des caractères spéciaux font partie du même groupe, prend les aussi en compte
        (ex: Autres et <Autres> font partie du même groupe, ou encore Autres et autres, ou bien Autre et AUTRES)
        
        RÉPONSE: Renvoie UNIQUEMENT une liste JSON de listes, où chaque sous-liste contient les noms des projets à regrouper ensemble.
        Si un projet n'a pas de similaire, ne l'inclus pas dans la réponse.
        
        Exemple de format de réponse:
        [
            ["NU", "NU Gate"],
            ["NU", "NU_MAX"],
            ["Projet1", "Projet1_Sub"]
        ]
        
        RÉPONDS UNIQUEMENT AVEC LA LISTE JSON, SANS AUTRE TEXTE."""
        
        return self.invoke_small_model(prompt)
    
    def generate_introduction(self, system_prompt: str) -> str:
        """
        Generate an introduction for a set of PPTX files.
        
        Args:
            system_prompt (str): The system prompt containing file information
            
        Returns:
            str: Generated introduction
        """
        prompt = f"""Tu es un assistant qui va générer une introduction pour un ensemble de fichiers PPTX. 
        Je veux juste une description globale des fichiers impliqués dans le message de l'utilisateur, 
        pas de cas par cas et surtout quelque chose de concis. 
        Renvoie uniquement l'introduction (pas d'explication). 
        Si tu vois une information importante ou une alerte critique, tu dois la signaler dans l'introduction. 
        
        Voici le contenu de tous les fichiers : {system_prompt} 
        
        Tu dois renvoyer uniquement l'introduction (pas d'explication)."""
        
        return self.invoke_small_model(prompt)

# Global model manager instance
model_manager = ModelManager() 