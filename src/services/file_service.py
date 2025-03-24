from pathlib import Path

def get_files_from_folder(folder_path: str) -> list[Path]:
    """
    Get paths of all files from a specified folder.
    
    Args:
        folder_path (str): Path to the folder to scan
        
    Returns:
        list[Path]: List of Path objects for each file in the folder
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        return []
    
    # Get all files in the folder (non-recursive)
    files = [f for f in folder.iterdir() if f.is_file()]
    return files

def get_files_from_folder_recursive(folder_path: str) -> list[Path]:
    """
    Get paths of all files from a specified folder and its subfolders.
    
    Args:
        folder_path (str): Path to the folder to scan
        
    Returns:
        list[Path]: List of Path objects for each file in the folder and subfolders
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        return []
    
    # Get all files in the folder and subfolders (recursive)
    files = [f for f in folder.rglob('*') if f.is_file()]
    return files

