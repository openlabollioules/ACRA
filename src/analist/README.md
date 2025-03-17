# PowerPoint Project Information Extractor

This set of scripts extracts project information from PowerPoint tables and formats it into a structured JSON output. It's specifically designed to handle tables in the first slide of PowerPoint presentations that contain project information.

## Features

- Extract project information from tables in PowerPoint slides
- Identify project names (in bold and underlined text)
- Associate information with the corresponding project
- Detect color-coded text (green for advancements, orange for small alerts, red for critical alerts)
- Format the information into a structured JSON output
- Process single PowerPoint files or entire folders
- Generate human-readable summaries

## Table Formats Supported

The script is designed to handle two table formats:

1. Single-column format with 4 rows:
   ```
   Section Title
   Information field (project name, project information)
   Section Title
   Information field (project name, project information)
   ```

2. Two-column format with 2 rows:
   ```
   Section Title | Section Title
   Information field | Information field
   ```

## Installation

1. Ensure you have Python 3.6+ installed
2. Install required dependencies:
   ```
   pip install python-pptx
   ```

## Usage

### Basic Usage

```bash
python extract_project_info.py path/to/your/presentation.pptx
```

### Output to a JSON File

```bash
python extract_project_info.py path/to/your/presentation.pptx -o output.json
```

### Print a Human-Readable Summary

```bash
python extract_project_info.py path/to/your/presentation.pptx -s
```

### Process All PowerPoint Files in a Folder

```bash
python extract_project_info.py dummy.pptx -f path/to/folder/with/pptx/files
```

## JSON Output Format

The script generates a JSON output with the following structure:

```json
{
  "Project Name 1": {
    "information": "This is the information about Project Name 1. <rgb=0 150 0 >Great progress on the frontend development.<rgb=0 150 0 >",
    "alerts": {
      "advancements": ["Great progress on the frontend development."],
      "small_alerts": [],
      "critical_alerts": []
    }
  },
  "Project Name 2": {
    "information": "Project Name 2 information...",
    "alerts": {
      "advancements": [],
      "small_alerts": [],
      "critical_alerts": []
    }
  }
}
```

## How Project Information is Extracted

1. The script analyzes the first slide of the PowerPoint presentation
2. It locates tables in the slide
3. It processes the text in each table cell, looking for:
   - Project names (text that is both bold and underlined)
   - Information associated with each project (text between project names)
   - Color-coded text using RGB values

## Color Classification

- **Green** (high G value): Advancements/Progress
- **Orange** (high R and G): Small Alerts/Concerns
- **Red** (high R value): Critical Alerts/Issues

## Files in this Package

- `project_extractor.py`: Core functionality for extracting information from PowerPoint tables
- `project_json_formatter.py`: Formats extracted information into JSON and provides human-readable summaries
- `extract_project_info.py`: Main script that combines extraction and formatting functionality
- `example_output.json`: Example of the JSON output format

## Requirements

- Python 3.6+
- python-pptx library 