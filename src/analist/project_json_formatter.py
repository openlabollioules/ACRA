import json
import re
from typing import Dict, List, Any, Optional
from project_extractor import extract_and_format_projects

def analyze_rgb_tags(text: str) -> Dict[str, List[str]]:
    """
    Analyze text with RGB tags and extract color-coded portions.
    
    Text inside rgb tags is analyzed as follows:
    - Green (high G value): big advancement
    - Orange (high R and G): small alert
    - Red (high R value): critical alert
    """
    patterns = {
        # Matches <rgb=R G B >TEXT<rgb=R G B >
        "rgb": re.compile(r'<rgb=(\d+) (\d+) (\d+) >(.*?)<rgb=\1 \2 \3 >')
    }
    
    results = {
        "advancements": [],
        "small_alerts": [],
        "critical_alerts": [],
        "original_text": text
    }
    
    # Find all RGB tagged content
    rgb_matches = patterns["rgb"].findall(text)
    
    for r, g, b, content in rgb_matches:
        r, g, b = int(r), int(g), int(b)
        
        # Classify based on RGB values
        if g > max(r, b) + 50:  # Green dominant
            results["advancements"].append(content)
        elif r > g + 50 and g > b + 50:  # Orange-ish
            results["small_alerts"].append(content)
        elif r > max(g, b) + 50:  # Red dominant
            results["critical_alerts"].append(content)
    
    return results

def format_project_data(raw_data: Dict[str, Dict]) -> Dict[str, Dict]:
    """
    Format raw project data into a more structured and clean format.
    """
    formatted_data = {}
    
    for project_name, project_info in raw_data.items():
        # Skip empty project names
        if not project_name.strip():
            continue
            
        # Get information field
        info = project_info.get("information", "")
        
        # Analyze RGB tags in the information
        rgb_analysis = analyze_rgb_tags(info)
        
        # Format the project data
        formatted_data[project_name] = {
            "information": info,
            "alerts": {
                "advancements": rgb_analysis["advancements"],
                "small_alerts": rgb_analysis["small_alerts"],
                "critical_alerts": rgb_analysis["critical_alerts"]
            }
        }
    
    return formatted_data

def extract_and_format_json_output(pptx_file: str, output_file: Optional[str] = None) -> Dict[str, Any]:
    """
    Process a PowerPoint file, extract project information, and format it nicely as JSON.
    """
    # Extract raw project data from the presentation
    raw_data = extract_and_format_projects(pptx_file)
    
    # Format the raw data into a cleaner structure
    formatted_data = format_project_data(raw_data)
    
    # Output the formatted data as JSON
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(formatted_data, f, indent=2, ensure_ascii=False)
    
    return formatted_data

def print_project_summary(project_data: Dict[str, Dict]) -> None:
    """
    Print a human-readable summary of the project data.
    """
    if not project_data:
        print("No project data found.")
        return
    
    print("\n=== PROJECT SUMMARY ===\n")
    
    for project_name, data in project_data.items():
        print(f"PROJECT: {project_name}")
        print("-" * (len(project_name) + 9))
        
        # Print information with highlighted alerts
        print("\nINFORMATION:")
        print(data["information"].replace("<rgb=", "[").replace(">", "]"))
        
        # Print advancements (green)
        advancements = data["alerts"]["advancements"]
        if advancements:
            print("\nADVANCEMENTS:")
            for i, item in enumerate(advancements, 1):
                print(f"  {i}. {item}")
        
        # Print small alerts (orange)
        small_alerts = data["alerts"]["small_alerts"]
        if small_alerts:
            print("\nSMALL ALERTS:")
            for i, item in enumerate(small_alerts, 1):
                print(f"  {i}. {item}")
        
        # Print critical alerts (red)
        critical_alerts = data["alerts"]["critical_alerts"]
        if critical_alerts:
            print("\nCRITICAL ALERTS:")
            for i, item in enumerate(critical_alerts, 1):
                print(f"  {i}. {item}")
        
        print("\n" + "=" * 40 + "\n")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pptx_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        # Process the file and get formatted data
        formatted_data = extract_and_format_json_output(pptx_file, output_file)
        
        # Print a human-readable summary
        print_project_summary(formatted_data)
        
        if output_file:
            print(f"\nJSON data has been saved to {output_file}")
    else:
        print("Usage: python project_json_formatter.py <pptx_file> [output_json_file]") 