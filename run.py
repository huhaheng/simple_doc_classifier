#!/usr/bin/env python3
"""
Main entry point for the Document Analysis for Windows
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
sys.path.insert(0, str(Path(__file__).parent / "src"))

try:
    from src.doc_classify_keywords import main
    
    if __name__ == "__main__":
        main()
        
except ImportError as e:
    print("Error: Missing required dependencies.")
    print("Please install the required packages:")
    print("pip install -r requirements.txt")
    print(f"Error details: {e}")
    sys.exit(1)
except KeyboardInterrupt:
    print("\nProgram interrupted by user.")
    sys.exit(0)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    sys.exit(1) 