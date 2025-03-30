import os
import json
import uuid
import platform
import subprocess
import time
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
            
        # Set default parameters
        self.default_model = "gpt-4"
        self.use_llm_chaining = False
    
    def generate_presentation_outline(self, prompt):
        """Use OpenAI to generate a structured presentation outline"""
        try:
            # Using the new client.chat.completions.create format for OpenAI >= 1.0.0
            response = client.chat.completions.create(
                model=self.default_model,
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
    
    def research_slide_content(self, slide_info, main_topic):
        """Research and enhance content for a specific slide"""
        try:
            # Prepare the prompt for research
            slide_title = slide_info["title"]
            current_content = " ".join(slide_info["content"])
            
            research_prompt = f"""
            I'm creating a presentation slide titled "{slide_title}" as part of a presentation about "{main_topic}".
            
            The current content for this slide is:
            {current_content}
            
            Please research this topic in depth and provide:
            1. 3-5 enhanced bullet points with specific facts, data, or examples
            2. Detailed presenter notes with background information and talking points
            3. Any relevant statistics or case studies that would strengthen this slide
            
            Format your response as a JSON object with these fields:
            {{
                "bullet_points": ["Point 1 with specific data", "Point 2 with example", ...],
                "presenter_notes": "Detailed background information and talking points for the presenter",
                "references": ["Optional reference 1", "Optional reference 2", ...]
            }}
            """
            
            response = client.chat.completions.create(
                model=self.default_model,
                messages=[
                    {"role": "system", "content": "You are a research assistant specializing in creating well-researched, factual presentation content."},
                    {"role": "user", "content": research_prompt}
                ],
                temperature=0.5
            )
            
            content = response.choices[0].message.content
            
            # Extract JSON from the response
            json_start = content.find('{')
            json_end = content.rfind('}') + 1
            json_str = content[json_start:json_end]
            
            # Parse JSON
            enhanced_content = json.loads(json_str)
            
            # Update the slide information
            updated_slide = slide_info.copy()
            updated_slide["content"] = enhanced_content["bullet_points"]
            
            # Combine original notes with new detailed notes
            original_notes = slide_info.get("notes", "")
            if original_notes:
                updated_slide["notes"] = f"{original_notes}\n\n{enhanced_content['presenter_notes']}"
            else:
                updated_slide["notes"] = enhanced_content["presenter_notes"]
            
            # Add references if available
            if "references" in enhanced_content and enhanced_content["references"]:
                references_text = "\n\nReferences:\n" + "\n".join(f"- {ref}" for ref in enhanced_content["references"])
                updated_slide["notes"] += references_text
            
            return updated_slide
            
        except Exception as e:
            print(f"Error researching slide content: {e}")
            # Return the original slide if research fails
            return slide_info
    
    def enhance_presentation_outline(self, outline, main_topic):
        """Enhance the presentation outline with researched content for each slide"""
        enhanced_outline = {
            "title": outline["title"],
            "slides": []
        }
        
        print("Researching and enhancing slide content...")
        total_slides = len(outline["slides"])
        
        for i, slide in enumerate(outline["slides"]):
            print(f"Researching slide {i+1}/{total_slides}: {slide['title']}...")
            enhanced_slide = self.research_slide_content(slide, main_topic)
            enhanced_outline["slides"].append(enhanced_slide)
            
            # Add a small delay to avoid rate limiting
            if i < total_slides - 1:
                time.sleep(0.5)
        
        return enhanced_outline
    
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
    
    def generate(self, prompt, convert_to_pdf=True, use_llm_chaining=False):
        """Generate a complete presentation from a prompt"""
        self.use_llm_chaining = use_llm_chaining
        
        # Step 1: Generate initial outline
        print("Step 1: Generating presentation outline...")
        outline = self.generate_presentation_outline(prompt)
        
        # Step 2 (Optional): If LLM chaining is enabled, research and enhance each slide
        if use_llm_chaining:
            print("Step 2: Researching and enhancing slide content...")
            outline = self.enhance_presentation_outline(outline, prompt)
        
        # Step 3: Create the PowerPoint presentation
        print("Step 3: Creating PowerPoint presentation...")
        pptx_path = self.create_presentation(outline)
        
        # Step 4: Convert to PDF if requested
        pdf_path = None
        if convert_to_pdf:
            print("Step 4: Converting to PDF...")
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
    result = generator.generate(prompt, use_llm_chaining=True)
    
    print(f"PowerPoint created: {result['pptx_path']}")
    if result['pdf_path']:
        print(f"PDF created: {result['pdf_path']}")
    else:
        print("PDF conversion not successful. You may need to install LibreOffice or PowerPoint.") 