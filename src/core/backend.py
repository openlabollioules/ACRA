import os
from dotenv import load_dotenv
import sys
from typing import Optional, Dict, Any # Added for type hinting
import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from services import update_table_with_project_data
from analist import analyze_presentation_with_colors, extract_projects_from_presentation
from .extract_and_summarize import aggregate_and_summarize, Generate_pptx_from_text

load_dotenv()
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "pptx_folder")
if not os.path.isabs(UPLOAD_FOLDER):
    UPLOAD_FOLDER = os.path.join(BASE_DIR, UPLOAD_FOLDER)

OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "OUTPUT")
if not os.path.isabs(OUTPUT_FOLDER):
    OUTPUT_FOLDER = os.path.join(BASE_DIR, OUTPUT_FOLDER)

def summarize_ppt(chat_id: str, add_info: Optional[str] = None, timestamp: Optional[str] = None, raw_structure_data: Optional[Dict[str, Any]] = None):
    """
    Summarizes content from PowerPoint files for a given chat_id or uses provided raw_structure_data.
    Then, updates a template PowerPoint file with this summarized data.

    This function orchestrates the entire summarization workflow:
    1. Aggregates and summarizes data from PowerPoint files or uses provided structure
    2. Validates the structure to ensure it contains required data
    3. Prepares the data for PowerPoint generation
    4. Creates a PowerPoint presentation using a template
    5. Handles errors at each step with appropriate fallbacks

    Args:
        chat_id (str): The identifier for the conversation (previously folder_name).
        add_info (str, optional): Additional information to include in the summary prompt.
        timestamp (str, optional): Timestamp for unique filenames. Auto-generated if None.
        raw_structure_data (dict, optional): Pre-extracted project structure. If provided, file aggregation is skipped by aggregate_and_summarize.

    Returns:
        dict: Contains the filename and path to the summarized PowerPoint file, or an error structure.
              Format: {"filename": str, "summary": str} or {"error": str, "summary": None}
    """
    
    # Log the start of the process
    print(f"Starting summarization for chat_id: {chat_id}")
    if raw_structure_data:
        print("summarize_ppt received raw_structure_data, will pass to aggregate_and_summarize.")

    # Step 1: Aggregate and summarize the data
    # This function either processes files in the chat_id's folder or uses provided raw_structure_data
    # It returns a JSON structure with projects, upcoming events, and metadata
    summarized_json_structure = aggregate_and_summarize(
        chat_id=chat_id, 
        add_info=add_info,
        timestamp=timestamp,
        raw_structure_data=raw_structure_data
    )

    # Step 2: Validate the structure returned by aggregate_and_summarize
    # If the structure is invalid, return an error
    if not isinstance(summarized_json_structure, dict) or "projects" not in summarized_json_structure:
        # Build a detailed error message with available information
        error_detail = summarized_json_structure.get("error", "Invalid structure from summarization") if isinstance(summarized_json_structure, dict) else "Unexpected response from summarization"
        log_message = f"ERROR: {error_detail} for chat_id {chat_id}."
        if isinstance(summarized_json_structure, dict) and summarized_json_structure.get("metadata", {}).get("errors"):
            log_message += f" Details: {summarized_json_structure['metadata']['errors']}"
        print(log_message)
        # Return an error structure compatible with CommandHandler expectations
        return {"error": log_message, "summary": None} 

    # Step 3: Extract necessary data for PowerPoint generation
    # Extract project data and upcoming events
    project_data_for_pptx = summarized_json_structure.get("projects", {})
    upcoming_events_for_pptx = summarized_json_structure.get("upcoming_events", {})
    
    # Step 4: Check if we have any actual data to put in the PowerPoint
    # The LLM might return an empty "projects" dict if it couldn't summarize anything meaningful
    if not project_data_for_pptx and not upcoming_events_for_pptx: # If both are empty
        # Check if there were errors during the summarization process
        metadata_errors = summarized_json_structure.get("metadata", {}).get("errors", [])
        source_file_errors = [sf.get("error") for sf in summarized_json_structure.get("source_files", []) if sf.get("error")]
        all_errors = metadata_errors + source_file_errors

        if all_errors:
            error_message = f"No project data to populate PowerPoint for chat {chat_id}. Errors encountered: {'; '.join(all_errors)}"
        else:
            # This might be valid if input files were empty or LLM deemed nothing summarizable
            error_message = f"No summarizable project data or upcoming events found to populate PowerPoint for chat {chat_id}. The input might have been empty or non-relevant."
        
        print(f"WARNING: {error_message}")
        # We still proceed to generate a potentially empty PowerPoint file

    # Step 5: PowerPoint Generation
    # Create output directory for this chat_id
    chat_summary_output_dir = os.path.join(OUTPUT_FOLDER, chat_id, "summaries")
    os.makedirs(chat_summary_output_dir, exist_ok=True)
    
    # Generate timestamp if not provided by caller
    current_timestamp = timestamp if timestamp else datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Define the output file path
    output_filename = os.path.join(chat_summary_output_dir, f"summary_{chat_id}_{current_timestamp}.pptx")
    
    print(f"Creating summary PowerPoint at: {output_filename} for chat_id: {chat_id}")
    
    # Step 6: Get the template file path
    template_path = os.getenv("TEMPLATE_FILE", "templates/CRA_TEMPLATE_IA.pptx")
    if not os.path.isabs(template_path):
        template_path = os.path.join(BASE_DIR, template_path)

    if not os.path.exists(template_path):
        print(f"WARNING: Template file not found at {template_path}. update_table_with_project_data might fail or use a default.")
        # update_table_with_project_data should ideally handle template absence gracefully or take a Presentation object.

    # Call update_table_with_project_data to create/update the PowerPoint
    try:
        final_ppt_path = update_table_with_project_data(
            pptx_path=template_path, 
            slide_index=0,      # Assuming first slide
            table_shape_index=0,      # Assuming first table on that slide
            project_data=project_data_for_pptx,
            output_path=output_filename,
            upcoming_events=upcoming_events_for_pptx
        )
        if "error" in final_ppt_path.lower(): # If update_table_with_project_data returns an error string
            print(f"Error from update_table_with_project_data for chat {chat_id}: {final_ppt_path}")
            return {"error": f"PPTX generation failed: {final_ppt_path}", "summary": None}

        created_filename = os.path.basename(final_ppt_path)
        print(f"Successfully created summary PowerPoint: {final_ppt_path} for chat_id: {chat_id}")
        return {"filename": created_filename, "summary": final_ppt_path} # Return path for upload

    except Exception as e:
        error_msg = f"Exception during PowerPoint generation for chat {chat_id}: {str(e)}"
        print(f"ERROR: {error_msg}")
        import traceback
        traceback.print_exc() # Print full traceback for debugging
        return {"error": error_msg, "summary": None}

def get_slide_structure(foldername : str):
    """
    Analyzes all PowerPoint presentations in a folder and extracts their structured content.
    
    This function:
    1. Finds all PowerPoint files in the specified folder
    2. Extracts project data and upcoming events from each file
    3. Merges the data from all files into a coherent structure
    4. Tracks metadata about the processing for reporting
    
    The resulting structure is used for summarization, analysis, and PowerPoint generation.
    
    Args:
        foldername (str): Name of the folder containing PowerPoint files to analyze
        
    Returns:
        dict: A structured representation of all PowerPoint content with:
              - projects: Hierarchical project data
              - upcoming_events: Events organized by service
              - metadata: Processing information
              - source_files: Details about processed files
              
    Raises:
        Exception: If the folder doesn't exist or other processing errors occur
    """
    # Check if foldername is None
    if foldername is None:
        raise Exception("Le nom du dossier (foldername) ne peut pas être None.")
        
    # Build the full path to the folder and validate it exists
    folder_path = os.path.join(UPLOAD_FOLDER, foldername)
    if not os.path.exists(folder_path):
        raise Exception("Le dossier n'existe pas.")

    # Find all PowerPoint files in the folder
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith(".pptx")]

    # Handle the case where no PowerPoint files are found
    if not pptx_files:
        return {"message": "Aucun fichier PPTX fourni."}
    
    # ===== HELPER FUNCTIONS =====
    
    def merge_project_dictionaries(dict1, dict2):
        """
        Recursively merges two project dictionaries, handling nested structures.
        
        This handles both terminal nodes (with 'information' field) and 
        intermediate nodes (with nested project dictionaries).
        
        Args:
            dict1 (dict): First dictionary (base)
            dict2 (dict): Second dictionary (to merge into base)
            
        Returns:
            dict: Merged dictionary
        """
        result = dict1.copy()
        
        for key, value in dict2.items():
            if key in result:
                # If the key exists in both dictionaries
                if isinstance(value, dict) and isinstance(result[key], dict):
                    # If both values are dictionaries, merge recursively
                    if "information" in value and "information" in result[key]:
                        # Terminal node - merge content fields
                        result[key]["information"] += "\n\n" + value["information"] if result[key]["information"] else value["information"]
                        result[key]["critical"].extend([item for item in value.get("critical", []) if item not in result[key]["critical"]])
                        result[key]["small"].extend([item for item in value.get("small", []) if item not in result[key]["small"]])
                        result[key]["advancements"].extend([item for item in value.get("advancements", []) if item not in result[key]["advancements"]])
                    else:
                        # Intermediate node - merge recursively
                        result[key] = merge_project_dictionaries(result[key], value)
                else:
                    # If types differ, keep the value from dict2
                    result[key] = value
            else:
                # If key doesn't exist in dict1, add it
                result[key] = value
                
        return result
    
    def extract_service_name(filename):
        """
        Extracts a readable service name from a PowerPoint filename.
        
        Args:
            filename (str): Filename of a PowerPoint file
            
        Returns:
            str: Extracted service name
        """
        # Extract title after the last underscore
        parts = filename.split('_')
        if len(parts) > 1:
            # Take parts after the first element (usually contains the ID)
            title_parts = parts[1:]
            # Recombine, removing .pptx extension
            title = ' '.join(title_parts).replace('.pptx', '').strip()
            # Properly capitalize
            title = ' '.join(word.capitalize() for word in title.split())
            return title
        else:
            # If no underscore, just remove the extension
            return filename.replace('.pptx', '').strip()

    # ===== MAIN PROCESSING =====
    
    # Initialize data structures to hold merged results
    all_projects = {}
    upcoming_events_by_service = {}
    processed_files = []
    
    # Process each PowerPoint file
    for filename in pptx_files:
        file_path = os.path.join(folder_path, filename)
        
        try:            
            # Extract project data using the extraction module
            project_data = extract_projects_from_presentation(file_path)
            
            # Extract service name from filename for categorization
            service_name = extract_service_name(filename)
            
            # Get upcoming events from metadata if available
            collected_events = []
            if "metadata" in project_data and "collected_upcoming_events" in project_data["metadata"]:
                collected_events = project_data["metadata"]["collected_upcoming_events"]
            
            # Add events to the corresponding service
            if collected_events:
                if service_name not in upcoming_events_by_service:
                    upcoming_events_by_service[service_name] = []
                for event in collected_events:
                    if event not in upcoming_events_by_service[service_name]:
                        upcoming_events_by_service[service_name].append(event)
            
            # Merge project data with existing projects
            if "projects" in project_data:
                all_projects = merge_project_dictionaries(all_projects, project_data["projects"])
            
            # Track successfully processed files
            processed_files.append({
                "filename": filename,
                "service_name": service_name,
                "processed": True,
                "events_count": len(collected_events)
            })
        except Exception as e:
            # Track files that couldn't be processed
            processed_files.append({"filename": filename, "error": f"Erreur lors de l'analyse: {str(e)}"})
    
    # Add default events message if no events were found for processed services
    if not upcoming_events_by_service:
        for file_info in processed_files:
            if "service_name" in file_info and "processed" in file_info and file_info["processed"]:
                service_name = file_info["service_name"]
                if service_name not in upcoming_events_by_service:
                    upcoming_events_by_service[service_name] = ["Aucun événement particulier prévu pour ce service."]
    
    # Create and return the final structure with all extracted data
    return {
        "projects": all_projects,
        "upcoming_events": upcoming_events_by_service,
        "metadata": {
            "processed_files": len(processed_files),
            "folder": foldername
        },
        "source_files": processed_files
    }

def get_slide_structure_wcolor(filename : str):
    """
    Analyzes a single PowerPoint file with color extraction.
    
    This function uses a different analysis method that preserves color information
    from the slides, which can be useful for certain visualization or analysis tasks.
    
    Args:
        filename (str): Name of the PowerPoint file to analyze
        
    Returns:
        dict: Analysis results with color information
        
    Raises:
        Exception: If the file is not found
    """
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    if not os.path.exists(file_path):
        raise Exception("File not found")

    # Analyze the presentation with color extraction
    slides_data = analyze_presentation_with_colors(file_path)
    return {"filename": filename, "slide data": slides_data}

def delete_all_pptx_files(foldername : str):
    """
    Deletes all PowerPoint files in the specified folder.
    
    This function is used for cleanup operations when files are no longer needed.
    
    Args:
        foldername (str): Name of the folder containing files to delete
        
    Returns:
        dict: Message indicating the result of the operation
        
    Raises:
        Exception: If the folder doesn't exist or files can't be deleted
    """
    pptx_folder = os.path.join(UPLOAD_FOLDER, foldername)
    if not os.path.exists(pptx_folder):
        raise Exception("Le dossier pptx_folder n'existe pas.")

    # List all files in the folder
    files = os.listdir(pptx_folder)
    
    if not files:
        return {"message": "Aucun fichier à supprimer."}

    # Delete files one by one
    for file in files:
        file_path = os.path.join(pptx_folder, file)
        try:
            os.remove(file_path)
        except Exception as e:
            raise Exception(f"Erreur lors de la suppression de {file}: {str(e)}")

    return {"message": f"{len(files)} fichiers supprimés avec succès."}

def generate_pptx_from_text(foldername : str, info : str, timestamp : str = None):
    """
    Takes textual information and generates a PowerPoint presentation from it.
    
    This function:
    1. Uses an LLM to extract structured data from raw text
    2. Organizes the data into projects, advancements, alerts, etc.
    3. Creates a PowerPoint presentation using the structured data
    
    Unlike summarize_ppt, this function doesn't require existing PowerPoint files -
    it creates content directly from text, making it useful for quick generation
    of presentations from meeting notes, emails, etc.
    
    Args:
        foldername (str): Name of the folder to store the PowerPoint
        info (str): Text to analyze for the PowerPoint generation
        timestamp (str, optional): Timestamp to use in the filename for uniqueness. 
                                   If None, will use the current time.
    
    Returns:
        dict: A dictionary containing filename and path to the generated PowerPoint
              Format: {"filename": str, "summary": str}
    """
    
    # Determine the target folder
    target_folder = UPLOAD_FOLDER
    if foldername:
        target_folder = os.path.join(UPLOAD_FOLDER, foldername)
    
    # Ensure the upload directory exists
    os.makedirs(target_folder, exist_ok=True)
    
    # Generate the structured project data from text using our LLM-based function
    project_data = Generate_pptx_from_text(foldername, info)
    print("project_data :", project_data)

    # Ensure project_data is properly structured
    if isinstance(project_data, dict):
        # Extract the projects and upcoming events
        if 'projects' in project_data:
            projects = project_data.get('projects', {})
            upcoming = project_data.get('upcoming_events', {})
        else:
            projects = project_data
            upcoming = {}
    else:
        # Handle unexpected format
        print("Warning: project_data is not a dictionary. Creating empty structure.")
        projects = {}
        upcoming = {}
    
    # Generate timestamp if not provided
    if timestamp is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create output filename in the target folder
    generated_filename = f"{foldername}_text_summary_{timestamp}.pptx"
    output_filename = os.path.join(target_folder, generated_filename)
    print(f"Creating text-generated PowerPoint at: {output_filename}")
    
    # Get template path from environment variables
    load_dotenv()
    template_path = os.getenv("TEMPLATE_FILE", "templates/CRA_TEMPLATE_IA.pptx")
    if not os.path.isabs(template_path):
        template_path = os.path.join(BASE_DIR, template_path)
    
    # Generate the PowerPoint using the template and structured data
    generated_pptx = update_table_with_project_data(
        pptx_path=template_path,
        slide_index=0,  # first slide
        table_shape_index=0,  # index of the table
        project_data=projects,
        output_path=output_filename,
        upcoming_events=upcoming
    )

    # Return the filename and path
    filename = os.path.basename(generated_pptx)
    return {"filename": filename, "summary": generated_pptx}

