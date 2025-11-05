# Certificate Generation System

A professional Python-based system for automatically generating certificates by overlaying participant information onto certificate templates. Supports batch processing from Excel files and outputs both PNG and PDF formats.

## Features

- **Batch Processing**: Generate hundreds of certificates from a single Excel file
- **Multiple Formats**: Outputs both PNG and PDF certificates
- **Flexible Templates**: Support for different certificate templates
- **Professional Logging**: Comprehensive logging for tracking and debugging
- **Error Handling**: Robust error handling with detailed error messages
- **Type Safety**: Full type hints for better code reliability
- **Configurable**: Easy-to-modify configuration system
- **Anti-aliased Text**: High-quality text rendering

## Project Structure

```
Certificate_Generation/
├── starting_ending.py        # Full certificate generator (name + dates)
├── juststart.py              # Simple certificate generator (name only)
├── config.py                 # Configuration management
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── certificate_prospects/    # Excel files with participant data
│   ├── health_interns.xlsx
│   └── Interns Datails copy.xlsx
├── templates/                # Certificate template images
│   ├── template1.png
│   └── template1-1.png
└── generated_certificates/   # Output directory (auto-created)
```

## Requirements

- Python 3.8 or higher
- See `requirements.txt` for Python package dependencies

## Installation

### 1. Clone the Repository

```bash
git clone <repository-url>
cd Certificate_Generation
```

### 2. Create Virtual Environment (Recommended)

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Linux/Mac:
source venv/bin/activate
# On Windows:
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Simple Certificate Generation (Name Only)

Generates certificates with just participant names:

```bash
python juststart.py
```

**Input Requirements:**
- Excel file with a `Name` column
- Certificate template image

### Full Certificate Generation (Name + Dates)

Generates certificates with participant names and internship dates:

```bash
python starting_ending.py
```

**Input Requirements:**
- Excel file with columns: `Name`, `from`, `to`, `s_month`, `e_month`, `s_year`, `e_year`
- Certificate template image

### Configuration

Edit `config.py` to customize:

```python
# File paths
PARTICIPANTS_FILE = "certificate_prospects/health_interns.xlsx"
TEMPLATE_FILE = "templates/template1.png"
OUTPUT_DIRECTORY = "generated_certificates/"

# Font settings
FONT_COLOR = (68, 102, 253)  # BGR format
FONT_SIZE = 3.8

# Text positioning
COORDINATE_X_ADJUSTMENT = 7
COORDINATE_Y_ADJUSTMENT = 78
```

## Input File Format

### Excel File Structure

Your Excel file should contain the following columns:

**Simple Certificate (juststart.py):**
| Name |
|------|
| John Doe |
| Jane Smith |

**Full Certificate (starting_ending.py):**
| Name | from | s_month | s_year | to | e_month | e_year |
|------|------|---------|--------|-------|---------|--------|
| John Doe | 25 | Mar | 22 | 14 | Apr | 22 |
| Jane Smith | 1 | Jan | 23 | 31 | Mar | 23 |

## Template Requirements

Certificate templates should be:
- PNG format
- High resolution (recommended: 1920x1080 or higher)
- Leave space in the center for text overlay
- Use professional design elements

## Output

The system generates:
- **PNG files**: High-quality image certificates
- **PDF files**: Print-ready PDF certificates (starting_ending.py only)
- **Log file**: `certificate_generation.log` for tracking

Output naming format:
- Simple: `{Name}.png`
- Full: `Internship_Completion_Certificate_{Name}.png` and `.pdf`

## Logging

The system provides detailed logging:

```
2025-11-05 10:30:45 - INFO - Certificate Generation System Started
2025-11-05 10:30:45 - INFO - Loading participants from: certificate_prospects/health_interns.xlsx
2025-11-05 10:30:45 - INFO - Found 10 participants to process
2025-11-05 10:30:46 - INFO - Generating certificate for: John Doe
2025-11-05 10:30:46 - INFO - Saved PNG: generated_certificates/Internship_Completion_Certificate_John Doe.png
2025-11-05 10:30:46 - INFO - Saved PDF: generated_certificates/Internship_Completion_Certificate_John Doe.pdf
```

Logs are saved to:
- Console output (real-time)
- `certificate_generation.log` file

## Troubleshooting

### Common Issues

**1. File Not Found Error**
```
FileNotFoundError: Required file not found: templates/template1.png
```
**Solution**: Ensure template files exist in the `templates/` directory

**2. Missing Column Error**
```
KeyError: 'Name'
```
**Solution**: Verify your Excel file contains all required columns

**3. Empty Output Directory**
```
No certificates generated
```
**Solution**: Check the output directory exists and you have write permissions

**4. Image Not Loading**
```
Failed to load template
```
**Solution**: Ensure template is a valid PNG file and path is correct

### Getting Help

Check the log file `certificate_generation.log` for detailed error messages.

## Customization Guide

### Adjusting Text Position

1. Open `config.py`
2. Modify position values:
```python
name_x_offset: int = 7      # Move text left (negative) or right (positive)
name_y_offset: int = 78     # Move text up (positive) or down (negative)
```

### Changing Font Size

```python
name_font_size: float = 3.8  # Increase for larger text
date_font_size: float = 1.5  # Adjust date size independently
```

### Changing Font Color

```python
# Colors in BGR format (Blue, Green, Red)
font_color_bgr: Tuple[int, int, int] = (68, 102, 253)  # Professional blue
# Examples:
# (0, 0, 0) = Black
# (255, 255, 255) = White
# (0, 0, 255) = Red
```

## Code Quality

This codebase follows professional Python standards:

- **PEP 8 Compliance**: Consistent code style
- **Type Hints**: Full type annotations for reliability
- **Docstrings**: Comprehensive documentation for all functions
- **Error Handling**: Robust exception handling
- **Logging**: Professional logging system
- **Modular Design**: Clean separation of concerns
- **Configuration Management**: Centralized settings

## Performance

- Processes ~100 certificates per minute (typical)
- Memory efficient (processes one certificate at a time)
- Automatic resource cleanup

## Contributing

When contributing, please:
1. Follow PEP 8 style guidelines
2. Add type hints to all functions
3. Include docstrings for new functions
4. Add appropriate error handling
5. Update tests and documentation

## License

[Add your license information here]

## Author

Certificate Generation System
Date: 2025

## Version History

- **v2.0** (2025-11-05): Professional refactor with improved documentation
  - Added comprehensive docstrings
  - Implemented type hints
  - Enhanced error handling
  - Added logging system
  - Created configuration management

- **v1.0**: Initial version

## Support

For issues, questions, or contributions, please [add contact information or issue tracker link]
