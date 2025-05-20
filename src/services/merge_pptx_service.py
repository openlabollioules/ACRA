import subprocess
import sys
import os

# Path to the esbuild bundle
js_bundle = './src/services/dist/bundle.js'

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
        
        proc = subprocess.run(
            ['node', js_bundle, folder_path, output_path],
            capture_output=True,
            text=True
        )

        print('STDOUT:', proc.stdout)
        print('STDERR:', proc.stderr, file=sys.stderr)

        if proc.returncode != 0:
            return {
                "error": f"Merge failed: {proc.stderr}"
            }

        # The merged file should be in the output_path directory
        # Look for the most recent .pptx file
        merged_files = [f for f in os.listdir(output_path) if f.endswith('.pptx')]
        if not merged_files:
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
        return {
            "error": f"Error during merge: {str(e)}"
        }
