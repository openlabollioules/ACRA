import subprocess
import sys

# Path to the esbuild bundle
js_bundle = './src/services/dist/bundle.js'

def merge_pptx(folder_path: str, output_path: str):
    """
    Merge the pptx files in the folder into a single pptx file
    Args:
        folder_path: str, the path to the folder containing the pptx files to merge
        output_path: str, the path to the output file
    Returns:
        stdout: str, the stdout of the command
    """
    proc = subprocess.run(
        ['node', js_bundle, folder_path, output_path],  # you can pass args
        capture_output=True, text=True
    )

    print('STDOUT:', proc.stdout)
    print('STDERR:', proc.stderr, file=sys.stderr)

    return proc.stdout, proc.stderr
