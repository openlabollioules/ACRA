"""
ACRA Configuration Manager
Centralized configuration for the ACRA pipeline system
"""
import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))    
from OLLibrary.config.config_manager import Config

class ACRAConfig(Config):
    """
    ACRA-specific configuration manager that handles all environment variables
    and application constants in a centralized way.
    """
    
    def __init__(self, config_dir: str = None):
        super().__init__(config_dir)
        self._load_environment_defaults()
    
    def _load_environment_defaults(self):
        """Load default values from environment variables"""
        # Folder paths
        self.set_default("UPLOAD_FOLDER", os.getenv("UPLOAD_FOLDER", "pptx_folder"))
        self.set_default("OUTPUT_FOLDER", os.getenv("OUTPUT_FOLDER", "OUTPUT"))
        self.set_default("MAPPINGS_FOLDER", os.getenv("MAPPINGS_FOLDER", os.path.join(os.getcwd(), "mappings")))
        self.set_default("TEMPLATES_FOLDER", "templates")
        
        # API Configuration
        self.set_default("USE_API", self._str_to_bool(os.getenv("USE_API", "False")))
        self.set_default("API_URL", os.getenv("API_URL", "http://host.docker.internal:5050"))
        self.set_default("OPENWEBUI_API_URL", os.getenv("OPENWEBUI_API_URL", "http://host.docker.internal:3030/api/v1/"))
        self.set_default("OPENWEBUI_API_KEY", os.getenv("OPENWEBUI_API_KEY"))
        
        # Database paths
        self.set_default("OPENWEBUI_DB_PATH", os.getenv("OPENWEBUI_DB_PATH", "./open-webui/webui.db"))
        self.set_default("OPENWEBUI_UPLOADS", os.getenv("OPENWEBUI_UPLOADS", "open-webui/uploads"))
        
        # Model Configuration
        self.set_default("OLLAMA_BASE_URL", os.getenv("OLLAMA_BASE_URL", "http://host.docker.internal:11434"))
        self.set_default("STREAMING_MODEL", os.getenv("STREAMING_MODEL", "qwen3:30b-a3b"))
        self.set_default("SMALL_MODEL", os.getenv("SMALL_MODEL", "qwen2.5:14b"))
        self.set_default("MODEL_CONTEXT_SIZE", int(os.getenv("MODEL_CONTEXT_SIZE", "32000")))
        self.set_default("SMALL_MODEL_CONTEXT_SIZE", int(os.getenv("SMALL_MODEL_CONTEXT_SIZE", "16000")))
        
        # File processing
        self.set_default("MAX_FILE_SIZE_MB", int(os.getenv("MAX_FILE_SIZE_MB", "100")))
        self.set_default("ALLOWED_EXTENSIONS", [".pptx"])
        
        # Cleanup settings
        self.set_default("CLEANUP_RETENTION_HOURS", int(os.getenv("CLEANUP_RETENTION_HOURS", "24")))
        self.set_default("AUTO_CLEANUP_ENABLED", self._str_to_bool(os.getenv("AUTO_CLEANUP_ENABLED", "True")))
        
        # Template settings
        self.set_default("DEFAULT_TEMPLATE", "CRA_TEMPLATE_IA.pptx")
        
    def set_default(self, key: str, value):
        """Set a default value only if the key doesn't already exist"""
        if key not in self.config:
            self.config[key] = value
    
    def _str_to_bool(self, value: str) -> bool:
        """Convert string to boolean"""
        if isinstance(value, bool):
            return value
        return value.lower() in ("true", "1", "t", "yes", "y")
    
    # Convenience properties for commonly used paths
    @property
    def upload_folder(self) -> str:
        return os.path.abspath(self.get("UPLOAD_FOLDER"))
    
    @property
    def output_folder(self) -> str:
        return os.path.abspath(self.get("OUTPUT_FOLDER"))
    
    @property
    def mappings_folder(self) -> str:
        return os.path.abspath(self.get("MAPPINGS_FOLDER"))
    
    @property
    def templates_folder(self) -> str:
        return os.path.abspath(self.get("TEMPLATES_FOLDER"))
    
    @property
    def template_path(self) -> str:
        return os.path.join(self.templates_folder, self.get("DEFAULT_TEMPLATE"))
    
    def ensure_directories(self):
        """Ensure all required directories exist"""
        directories = [
            self.upload_folder,
            self.output_folder,
            self.mappings_folder,
            self.templates_folder
        ]
        
        for directory in directories:
            os.makedirs(directory, exist_ok=True)
    
    def get_conversation_upload_folder(self, chat_id: str) -> str:
        """Get the upload folder path for a specific conversation"""
        return os.path.join(self.upload_folder, chat_id)
    
    def get_conversation_output_folder(self, chat_id: str) -> str:
        """Get the output folder path for a specific conversation"""
        return os.path.join(self.output_folder, chat_id)
    
    def get_mapping_file_path(self, chat_id: str) -> str:
        """Get the mapping file path for a specific conversation"""
        return os.path.join(self.mappings_folder, f"{chat_id}_file_mappings.json")

# Global configuration instance
acra_config = ACRAConfig() 