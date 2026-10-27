"""Configuration module for loading environment variables and defining directory paths."""

import os
from pathlib import Path

from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Resolve absolute path based on project root
BASE_DIR = Path(__file__).resolve().parent

# Externalize paths with fallback defaults
INPUT_DIR = BASE_DIR / os.getenv("INPUT_DIR", "inputs")

# Retrieve user's Documents report root directory
BASE_DOCUMENTS_DIR = Path.home() / "Documents" / "日報"
BASE_DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)

# Set output directory for generated reports
OUTPUT_DIR = BASE_DOCUMENTS_DIR

TEMPLATE_PATH = BASE_DIR / os.getenv("TEMPLATE_PATH", "inputs/aspark_template.pptx")

# Automatically create required directories on startup
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)