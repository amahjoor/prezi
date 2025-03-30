#!/usr/bin/env python3
import os
import argparse
from dotenv import load_dotenv
from ppt_generator import PresentationGenerator

def main():
    """Command-line interface for the PowerPoint generator"""
    # Load environment variables
    load_dotenv()
    
    # Check if OPENAI_API_KEY is set
    if not os.getenv("OPENAI_API_KEY"):
        print("Error: OPENAI_API_KEY environment variable is not set.")
        print("Please set it in a .env file or export it in your shell.")
        return 1
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Generate PowerPoint presentations from natural language.')
    parser.add_argument('prompt', nargs='?', help='The prompt describing the presentation to generate')
    parser.add_argument('-o', '--output', help='Output file name (without extension)')
    parser.add_argument('-i', '--interactive', action='store_true', help='Interactive mode')
    parser.add_argument('--no-pdf', action='store_true', help='Disable PDF export')
    parser.add_argument('--research', action='store_true', help='Enable research mode with LLM chaining')
    
    args = parser.parse_args()
    
    # Create the generator
    generator = PresentationGenerator()
    
    if args.interactive or not args.prompt:
        # Interactive mode
        print("=== GPT-PowerPoint Generator ===")
        print("Enter a description of the presentation you want to create.")
        print("Type 'quit' or 'exit' to quit.")
        print()
        
        while True:
            prompt = input("Prompt> ")
            
            if prompt.lower() in ['quit', 'exit', 'q']:
                break
            
            if not prompt.strip():
                print("Please enter a prompt.")
                continue
            
            # Ask for PDF conversion preference
            want_pdf = input("Do you want to generate a PDF version as well? (Y/n): ").lower() != 'n'
            
            # Ask for research mode preference
            use_research = input("Do you want to use research mode for enhanced content? (y/N): ").lower() == 'y'
            if use_research:
                print("Research mode enabled. This will take longer but produce more detailed content.")
            
            print("Generating presentation...")
            try:
                result = generator.generate(prompt, convert_to_pdf=want_pdf, use_llm_chaining=use_research)
                
                print(f"PowerPoint created successfully: {result['pptx_path']}")
                
                if result['pdf_path']:
                    print(f"PDF created successfully: {result['pdf_path']}")
                elif want_pdf:
                    print("PDF creation was not successful. You may need to install LibreOffice or PowerPoint.")
            except Exception as e:
                print(f"Error: {str(e)}")
    else:
        # Single generation mode
        print("Generating presentation...")
        try:
            result = generator.generate(
                args.prompt, 
                convert_to_pdf=not args.no_pdf,
                use_llm_chaining=args.research
            )
            
            print(f"PowerPoint created successfully: {result['pptx_path']}")
            
            if result['pdf_path']:
                print(f"PDF created successfully: {result['pdf_path']}")
            elif not args.no_pdf:
                print("PDF creation was not successful. You may need to install LibreOffice or PowerPoint.")
        except Exception as e:
            print(f"Error: {str(e)}")
            return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 