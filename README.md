# GPT-PowerPoint Generator

A tool that generates PowerPoint presentations from natural language prompts using AI.

## Features

- Convert natural language descriptions into structured PowerPoint presentations
- Web interface for easy interaction
- API endpoint for programmatic access
- Customizable slide templates and themes

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Create a `.env` file with your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```
4. Run the application:
   ```
   python app.py
   ```
5. Open http://localhost:5000 in your browser

## Usage

1. Enter a description of the presentation you want to create
2. Click "Generate Presentation"
3. Download the generated PowerPoint file

## Project Structure

- `app.py`: Flask web application
- `ppt_generator.py`: Core presentation generation logic
- `templates/`: HTML templates for the web interface
- `static/`: CSS, JavaScript, and static assets
- `generated/`: Output directory for generated presentations 