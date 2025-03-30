# GPT-PowerPoint Generator

A tool that generates PowerPoint presentations from natural language prompts using AI.

## Features

- Convert natural language descriptions into structured PowerPoint presentations
- Web interface for easy interaction
- Command-line interface for scripting
- API endpoint for programmatic access
- PDF export functionality
- Docker support for easy deployment
- Makefile for common operations

## Requirements

### Python Dependencies
- Python 3.7 or higher
- Required packages installed via pip (see Setup section)

### External Dependencies (for PDF export)
To convert PowerPoint to PDF, one of the following is required:
- **macOS**: `brew install --cask libreoffice` or `brew install unoconv`
- **Linux**: Install LibreOffice or unoconv using your package manager
- **Windows**: Microsoft PowerPoint or LibreOffice

*Note: PDF export is optional. The application will still generate PowerPoint files without these dependencies.*

## Setup

### Standard Installation

1. Clone this repository
2. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Create a `.env` file with your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```
4. For PDF export, install the required external dependencies (see above)
5. Check your installation:
   ```
   python check_deps.py
   ```

### Using Docker

1. Install Docker and Docker Compose
2. Create a `.env` file with your OpenAI API key
3. Build and run the application:
   ```
   docker-compose up
   ```

## Running the Application

### Web Interface

```
python app.py
```
or
```
make run
```
Then open http://localhost:5000 in your browser.

### Command-Line Interface

For single presentation generation:
```
python cli.py "Create a presentation about quantum computing"
```

For interactive mode:
```
python cli.py -i
```

To skip PDF generation:
```
python cli.py --no-pdf "Create a presentation about quantum computing"
```

### Using Makefile Commands

- `make setup` - Install dependencies and create directories
- `make check` - Verify your installation
- `make run` - Start the web application
- `make cli` - Run the command-line interface
- `make clean` - Remove generated presentations and cache files
- `make test` - Generate a test presentation
- `make help` - Show all available commands

## Usage

### Web Interface
1. Enter a description of the presentation you want to create
2. Check the "Also generate PDF version" box if needed
3. Click "Generate Presentation"
4. Download the PowerPoint and/or PDF files

### API Endpoint
You can also generate presentations programmatically:
```
curl -X POST http://localhost:5000/api/generate \
  -H "Content-Type: application/json" \
  -d '{"prompt": "Create a presentation about AI ethics", "convert_to_pdf": true}'
```

## Project Structure

- `app.py`: Flask web application
- `ppt_generator.py`: Core presentation generation logic
- `cli.py`: Command-line interface
- `templates/`: HTML templates for the web interface
- `static/`: CSS, JavaScript, and static assets
- `generated/`: Output directory for generated presentations
- `Dockerfile` & `docker-compose.yml`: Docker configuration
- `Makefile`: Shortcuts for common commands
- `check_deps.py`: Utility to check dependencies 