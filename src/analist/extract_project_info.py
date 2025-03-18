#!/usr/bin/env python3
import os
import sys
import argparse
from project_extractor import extract_projects_from_presentation
from project_json_formatter import format_project_data, print_project_summary

def main():
    """
    Main entry point for the project information extractor and formatter.
    """
    parser = argparse.ArgumentParser(description='Extract project information from PowerPoint files.')
    parser.add_argument('pptx_file', help='Path to the PowerPoint file to analyze')
    parser.add_argument('-o', '--output', help='Path to output JSON file (optional)', default=None)
    parser.add_argument('-s', '--summary', action='store_true', help='Print a summary of extracted information')
    parser.add_argument('-f', '--folder', help='Process all PPTX files in the specified folder (instead of a single file)')
    
    args = parser.parse_args()
    
    # Check if processing a folder or a single file
    if args.folder:
        # Process all PPTX files in the folder
        if not os.path.isdir(args.folder):
            print(f"Error: {args.folder} is not a valid directory")
            return 1
        
        all_projects = {}
        
        for filename in os.listdir(args.folder):
            if filename.lower().endswith('.pptx'):
                file_path = os.path.join(args.folder, filename)
                print(f"Processing {filename}...")
                
                # Extract projects from this presentation
                projects = extract_projects_from_presentation(file_path)
                
                # Add to the combined results
                for project_name, info in projects.items():
                    if project_name in all_projects:
                        # Project already exists, append information
                        all_projects[project_name]["information"] += f"\n[From {filename}] {info['information']}"
                        
                        # Merge alerts
                        for alert_type in ["advancements", "small_alerts", "critical_alerts"]:
                            all_projects[project_name]["alerts"][alert_type].extend(
                                info["alerts"].get(alert_type, [])
                            )
                    else:
                        # New project
                        all_projects[project_name] = info
        
        # Format the combined data
        formatted_data = format_project_data(all_projects)
        
    else:
        # Process a single file
        if not os.path.isfile(args.pptx_file):
            print(f"Error: {args.pptx_file} is not a valid file")
            return 1
        
        # Extract projects from the presentation
        formatted_data = extract_projects_from_presentation(args.pptx_file)
        
        # Format the data
        # formatted_data = format_project_data(projects)
    
    # Output to JSON file if specified
    if args.output:
        import json
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(formatted_data, f, indent=2, ensure_ascii=False)
        print(f"Project information saved to {args.output}")
    
    # Print summary if requested
    if args.summary:
        print_project_summary(formatted_data)
    elif not args.output:
        # If no output specified and no summary requested, print the JSON to stdout
        import json
        print(json.dumps(formatted_data, indent=2, ensure_ascii=False))
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 