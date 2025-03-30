import os
import json
import uuid
import platform
import subprocess
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import openai
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure OpenAI client
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

class PresentationGenerator:
    def __init__(self):
        self.output_dir = "generated"
        
        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
    def generate_presentation_outline(self, prompt):
        """Use OpenAI to generate a structured presentation outline"""
        try:
            # Using the new client.chat.completions.create format for OpenAI >= 1.0.0
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": """You are an expert presentation designer. 
                    Create a detailed presentation outline with exactly the following JSON structure:
                    {
                        "title": "Main Presentation Title",
                        "slides": [
                            {
                                "title": "Slide Title",
                                "content": ["Bullet point 1", "Bullet point 2", "Bullet point 3"],
                                "notes": "Optional presenter notes"
                            },
                            ...more slides...
                        ]
                    }
                    
                    Follow these guidelines:
                    - Create 5-10 slides depending on the topic complexity
                    - Include a title slide, introduction, main content slides, and conclusion
                    - Keep bullet points concise and focused (1-2 lines each)
                    - Add presenter notes to provide additional context
                    - Ensure logical flow between slides
                    - ONLY output valid JSON that matches the structure above
                    """},
                    {"role": "user", "content": f"Generate a presentation outline about: {prompt}"}
                ],
                temperature=0.7
            )
            
            content = response.choices[0].message.content
            
            # Extract JSON from the response (in case there's surrounding text)
            json_start = content.find('{')
            json_end = content.rfind('}') + 1
            json_str = content[json_start:json_end]
            
            # Parse JSON
            outline = json.loads(json_str)
            return outline
            
        except Exception as e:
            print(f"Error generating outline: {e}")
            # Fallback outline if API fails
            return {
                "title": f"Presentation about {prompt}",
                "slides": [
                    {
                        "title": "Introduction", 
                        "content": ["Overview of the topic", "Key points to be covered"],
                        "notes": "Introduce yourself and the topic"
                    },
                    {
                        "title": "Key Points",
                        "content": ["Main idea 1", "Main idea 2", "Main idea 3"],
                        "notes": "Explain the main concepts"
                    },
                    {
                        "title": "Conclusion",
                        "content": ["Summary of key points", "Call to action or next steps"],
                        "notes": "Wrap up and invite questions"
                    }
                ]
            }
    
    def create_presentation(self, outline):
        """Create a PowerPoint presentation from the structured outline"""
        # Create a new presentation
        prs = Presentation()
        
        # Add title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = outline["title"]
        subtitle.text = "Generated Presentation"
        
        # Apply some basic styling to title slide
        title.text_frame.paragraphs[0].font.size = Pt(44)
        title.text_frame.paragraphs[0].font.bold = True
        
        # Add content slides
        for slide_info in outline["slides"]:
            content_slide_layout = prs.slide_layouts[1]  # Layout with title and content
            slide = prs.slides.add_slide(content_slide_layout)
            
            # Set slide title
            title = slide.shapes.title
            title.text = slide_info["title"]
            
            # Add content as bullet points
            content = slide.placeholders[1]
            text_frame = content.text_frame
            
            for i, bullet_point in enumerate(slide_info["content"]):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                p.text = bullet_point
                p.level = 0
            
            # Add notes if provided
            if "notes" in slide_info and slide_info["notes"]:
                slide.notes_slide.notes_text_frame.text = slide_info["notes"]
        
        # Generate a unique filename
        file_id = str(uuid.uuid4())[:8]
        file_name = f"{self.output_dir}/presentation_{file_id}.pptx"
        
        # Save the presentation
        prs.save(file_name)
        return file_name
    
    def convert_to_pdf(self, pptx_path):
        """Convert PowerPoint to PDF using appropriate method based on OS"""
        pdf_path = pptx_path.replace('.pptx', '.pdf')
        
        try:
            system = platform.system()
            
            if system == 'Darwin':  # macOS
                # Check if LibreOffice is installed
                try:
                    # Attempt to convert using LibreOffice
                    subprocess.run([
                        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                        '--headless',
                        '--convert-to', 'pdf',
                        '--outdir', os.path.dirname(pptx_path),
                        pptx_path
                    ], check=True)
                    print(f"PDF created with LibreOffice: {pdf_path}")
                    return pdf_path
                except (subprocess.SubprocessError, FileNotFoundError):
                    print("LibreOffice not found or error occurred.")
                    
                # Try using unoconv if LibreOffice direct method failed
                try:
                    subprocess.run(['unoconv', '-f', 'pdf', pptx_path], check=True)
                    print(f"PDF created with unoconv: {pdf_path}")
                    return pdf_path
                except (subprocess.SubprocessError, FileNotFoundError):
                    print("unoconv not found or error occurred.")
                    
            elif system == 'Windows':
                # On Windows, use comtypes to convert
                try:
                    import comtypes.client
                    
                    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
                    powerpoint.Visible = True
                    
                    deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path))
                    deck.SaveAs(os.path.abspath(pdf_path), 32)  # 32 is the PDF format code
                    deck.Close()
                    powerpoint.Quit()
                    
                    print(f"PDF created with PowerPoint: {pdf_path}")
                    return pdf_path
                except Exception as e:
                    print(f"Error using comtypes: {e}")
            
            elif system == 'Linux':
                # On Linux, use LibreOffice
                try:
                    subprocess.run([
                        'libreoffice', 
                        '--headless', 
                        '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(pptx_path),
                        pptx_path
                    ], check=True)
                    print(f"PDF created with LibreOffice: {pdf_path}")
                    return pdf_path
                except (subprocess.SubprocessError, FileNotFoundError):
                    print("LibreOffice not found or error occurred.")
                    
                # Try using unoconv as fallback
                try:
                    subprocess.run(['unoconv', '-f', 'pdf', pptx_path], check=True)
                    print(f"PDF created with unoconv: {pdf_path}")
                    return pdf_path
                except (subprocess.SubprocessError, FileNotFoundError):
                    print("unoconv not found or error occurred.")
            
            print(f"Could not convert to PDF on {system} system.")
            return None
            
        except Exception as e:
            print(f"Error during PDF conversion: {e}")
            return None
    
    def generate(self, prompt, convert_to_pdf=True):
        """Generate a complete presentation from a prompt"""
        # Generate PPT
        outline = self.generate_presentation_outline(prompt)
        pptx_path = self.create_presentation(outline)
        
        # Convert to PDF if requested
        pdf_path = None
        if convert_to_pdf:
            pdf_path = self.convert_to_pdf(pptx_path)
        
        # Return both paths in a dictionary
        return {
            'pptx_path': pptx_path,
            'pdf_path': pdf_path
        }


if __name__ == "__main__":
    # Test the generator
    generator = PresentationGenerator()
    prompt = "The future of artificial intelligence in healthcare"
    result = generator.generate(prompt)
    
    print(f"PowerPoint created: {result['pptx_path']}")
    if result['pdf_path']:
        print(f"PDF created: {result['pdf_path']}")
    else:
        print("PDF conversion not successful. You may need to install LibreOffice or PowerPoint.") 