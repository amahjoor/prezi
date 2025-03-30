.PHONY: setup check run clean test cli

setup:
	pip install -r requirements.txt
	mkdir -p generated

check:
	python check_deps.py

run:
	python app.py

cli:
	python cli.py

clean:
	rm -rf generated/*.pptx
	rm -rf __pycache__
	rm -rf *.pyc

test:
	python -c "from ppt_generator import PresentationGenerator; generator = PresentationGenerator(); generator.generate('Test presentation about AI technologies')"

help:
	@echo "Available commands:"
	@echo "  make setup  - Install dependencies and set up directories"
	@echo "  make check  - Check if dependencies are installed correctly"
	@echo "  make run    - Run the web application"
	@echo "  make cli    - Run the command-line interface"
	@echo "  make clean  - Remove generated presentations and cache files"
	@echo "  make test   - Generate a test presentation"
	@echo "  make help   - Show this help message" 