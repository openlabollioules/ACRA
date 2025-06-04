import subprocess
import sys
import os

from OLLibrary.utils.log_service import setup_logging, get_logger

# Path to the esbuild bundle
js_bundle = './src/services/dist/bundle.js'

setup_logging(app_name="Merge PPTX Service")
log = get_logger(__name__)

def merge_pptx(folder_path: str, output_path: str):
    """
    Merge the pptx files in the folder into a single pptx file
    Args:
        folder_path: str, the path to the folder containing the pptx files to merge
        output_path: str, the path to the output file
    Returns:
        dict: A dictionary containing:
            - merged_file: str, the path to the merged file if successful
            - error: str, error message if something went wrong
    """
    try:
        # Ensure output directory exists
        os.makedirs(output_path, exist_ok=True)

        log.info("Folder created. Launching merging process with node ...")
        
        proc = subprocess.run(
                ['node', js_bundle, folder_path, output_path],
                capture_output=True,
                text=True
        )

        log.info(f"Node process for merging ended. \nOut : {proc.stdout}.")

        print('STDOUT:', proc.stdout)
        print('STDERR:', proc.stderr, file=sys.stderr)

        if proc.returncode != 0:
            log.error(f"Merge failed : {proc.stderr}")
            return {
                "error": f"Merge failed: {proc.stderr}"
            }

        # The merged file should be in the output_path directory
        # Look for the most recent .pptx file
        merged_files = [f for f in os.listdir(output_path) if f.endswith('.pptx')]
        if not merged_files:
            log.error(f"No file created.")
            return {
                "error": "No merged file was created"
            }

        # Get the most recent file
        merged_file = max(
            [os.path.join(output_path, f) for f in merged_files],
            key=os.path.getctime
        )

        return {
            "merged_file": merged_file
        }

    except Exception as e:
        log.error(f"Error while merging : {str(e)}")
        return {
            "error": f"Error during merge: {str(e)}"
        }
