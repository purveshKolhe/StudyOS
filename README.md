# AI Presentation Generator

This project is a Flask-based web application that uses the Google Gemini API to automatically generate PowerPoint presentations from a given topic.

## Features

-   **Dynamic Content Generation**: Creates presentation slides based on a user-provided topic.
-   **Custom Templates**: Uses a base PowerPoint template for consistent styling and layout.
-   **Intelligent Layouts**: Leverages metadata to select appropriate slide layouts.
-   **Web Interface**: Simple and intuitive UI to enter a topic and download the generated presentation.
-   **7-10 Slides**: Generates a presentation with a variable number of slides, including mandatory slides like a title and thank you slide.

## Project Structure

```
.
├── app.py                  # Main Flask application logic
├── requirements.txt        # Python dependencies
├── generated/              # Output directory for generated presentations
├── static/                 # Frontend assets (CSS, JS)
│   ├── css/style.css
│   └── js/app.js
├── template/
│   ├── layout_metadata.json # Defines slide layouts and content rules
│   └── template.pptx        # Base presentation template
└── templates/
    └── index.html          # HTML for the web interface
```

## Setup and Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

3.  **Install the dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Set up your environment variables:**
    You need a Google Gemini API key. Create a `.env` file in the root directory and add your key:
    ```
    GEMINI_API_KEY="your_gemini_api_key"
    ```
    The application will load this key from the environment.

## Usage

1.  **Run the Flask application:**
    ```bash
    python app.py
    ```

2.  **Open your browser:**
    Navigate to `http://127.0.0.1:5000`.

3.  **Generate a presentation:**
    -   Enter a topic in the input field.
    -   Click the "Generate" button.
    -   Wait for the AI to create the presentation plan and build the `.pptx` file.
    -   Click the "Download" link to save your presentation.
