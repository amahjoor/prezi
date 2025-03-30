document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('presentation-form');
    const generateBtn = document.getElementById('generate-btn');
    const spinner = document.getElementById('spinner');
    const resultContainer = document.getElementById('result-container');
    const errorContainer = document.getElementById('error-container');
    const downloadPptxBtn = document.getElementById('download-pptx-btn');
    const downloadPdfBtn = document.getElementById('download-pdf-btn');
    const errorMessage = document.getElementById('error-message');
    const tryAgainBtn = document.getElementById('try-again-btn');
    const convertToPdfCheckbox = document.getElementById('convert_to_pdf');
    const useLlmChainingCheckbox = document.getElementById('use_llm_chaining');
    
    // Handle form submission
    form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        // Get form values
        const prompt = document.getElementById('prompt').value.trim();
        const convertToPdf = convertToPdfCheckbox.checked;
        const useLlmChaining = useLlmChainingCheckbox.checked;
        
        if (!prompt) {
            showError('Please enter a description for your presentation.');
            return;
        }
        
        // Show loading state
        setLoading(true);
        
        // If using research mode, update button text to indicate longer process
        if (useLlmChaining) {
            generateBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Researching and Generating (this may take a while...)';
        }
        
        // Send request to generate presentation
        fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: new URLSearchParams({
                'prompt': prompt,
                'convert_to_pdf': convertToPdf,
                'use_llm_chaining': useLlmChaining
            })
        })
        .then(response => {
            // Check for HTTP status codes first
            if (response.status === 429) {
                throw new Error('OpenAI API quota exceeded. Please update your API key or try again later.');
            }
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            if (data.success) {
                // Update download buttons with the correct file paths
                downloadPptxBtn.href = `/download/${data.pptx_filename}`;
                
                // Handle PDF download button if available
                if (data.has_pdf && data.pdf_filename) {
                    downloadPdfBtn.href = `/download/${data.pdf_filename}`;
                    downloadPdfBtn.classList.remove('d-none');
                } else {
                    downloadPdfBtn.classList.add('d-none');
                }
                
                // Show success container
                resultContainer.classList.remove('d-none');
                errorContainer.classList.add('d-none');
            } else {
                showError(data.error || 'An error occurred while generating the presentation.');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            
            // Display a user-friendly message for API quota errors
            if (error.message.includes('quota exceeded') || error.message.includes('insufficient_quota')) {
                showError('OpenAI API quota exceeded. The API key has reached its usage limit. Please update your API key or try again later.');
            } else {
                showError('An error occurred while connecting to the server. Please try again.');
            }
        })
        .finally(() => {
            setLoading(false);
            // Reset button text
            generateBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2 d-none" id="spinner" role="status" aria-hidden="true"></span>Generate Presentation';
        });
    });
    
    // Try again button handler
    tryAgainBtn.addEventListener('click', function() {
        errorContainer.classList.add('d-none');
        resultContainer.classList.add('d-none');
    });
    
    // Helper function to show error message
    function showError(message) {
        errorMessage.innerHTML = message;
        errorContainer.classList.remove('d-none');
        resultContainer.classList.add('d-none');
    }
    
    // Helper function to set loading state
    function setLoading(isLoading) {
        if (isLoading) {
            generateBtn.disabled = true;
            spinner.classList.remove('d-none');
            generateBtn.classList.add('btn-loading');
        } else {
            generateBtn.disabled = false;
            spinner.classList.add('d-none');
            generateBtn.classList.remove('btn-loading');
        }
    }
    
    // Add event listener for research mode checkbox
    useLlmChainingCheckbox.addEventListener('change', function() {
        const infoText = document.querySelector('.form-text');
        if (this.checked) {
            infoText.classList.add('text-primary');
            generateBtn.textContent = 'Generate Researched Presentation';
        } else {
            infoText.classList.remove('text-primary');
            generateBtn.textContent = 'Generate Presentation';
        }
    });
}); 