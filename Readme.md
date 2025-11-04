# PDF to PowerPoint Converter

A simple and effective Python script to convert PDF files into PowerPoint presentations. Each page of the PDF is converted into a separate slide in the generated `.pptx` file.

## Features

*   Converts each page of a PDF into a high-quality image.
*   Creates a new PowerPoint presentation from the converted images.
*   Each PDF page corresponds to one slide in the PowerPoint file.
*   Clean and easy-to-use script.

## Installation

To get started with this project, follow these steps:

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd Pdf-to-ppt
    ```

2.  **Create and activate a virtual environment:**
    This project uses a virtual environment to manage dependencies.

    *   On macOS and Linux:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```

    *   On Windows:
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```

3.  **Install dependencies:**
    The required Python libraries are listed in `requirements.txt`. Install them using pip:
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: If you don't have a `requirements.txt` file, you can create one by running `pip freeze > requirements.txt` after installing the necessary libraries like `python-pptx` and `PyMuPDF`)*

## Usage

To convert a PDF file, run the main script from the terminal.

```bash
python your_script_name.py path/to/your/file.pdf
