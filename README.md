# Clinton Pole BoM Generator

A Streamlit web application for generating Bills of Materials (BoM) for Clinton poles and accessories used in RFID installations.

## Features

- Easy-to-use web interface
- Input for Project ID and Reader Count
- Quantity selection for different pole types
- Automatic inclusion of accessories based on reader count
- CSV export functionality
- Clear validation messages

## Installation

1. Ensure Python 3.8+ is installed on your system
2. Clone or download this repository
3. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Running the Application

To run the application, navigate to the project directory in your terminal/command prompt and execute:

```bash
streamlit run clinton_bom_app.py
```

Your default web browser should automatically open with the application running. If not, look for a URL in the terminal output (usually http://localhost:8501) and open it manually.

## Adding More Parts

To add additional Clinton parts to the catalog:

1. Open `clinton_bom_app.py` in a text editor
2. Find the `clinton_parts` dictionary near the top of the file
3. Add new entries following the existing format:
   ```python
   "PART-NUMBER": {"desc": "Full Description", "type": "pole" or "accessory"}
   ```
4. Save the file and restart the application

## Customization

- The application uses a wide layout by default for better usability
- Column arrangement can be adjusted in the layout section
- CSV filename format can be modified in the download section

## Dependencies

- Streamlit
- Pandas

These are specified in the `requirements.txt` file.
