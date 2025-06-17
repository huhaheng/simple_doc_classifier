#!/usr/bin/env python3
"""
Setup verification script for Simple Document Analysis for Windows
"""

import sys
import os
from pathlib import Path

def check_python_version():
    """Check if Python version is compatible"""
    version = sys.version_info
    if version.major == 3 and version.minor >= 8:
        print(f"✓ Python {version.major}.{version.minor}.{version.micro} - Compatible")
        return True
    else:
        print(f"✗ Python {version.major}.{version.minor}.{version.micro} - Requires Python 3.8+")
        return False

def check_dependencies():
    """Check if all required dependencies are available"""
    dependencies = {
        'python-docx': 'docx',
        'python-docx2txt': 'docx2txt',
        'pywin32': 'win32com.client'
    }
    
    missing = []
    optional_missing = []
    
    for package_name, import_name in dependencies.items():
        try:
            __import__(import_name)
            print(f"✓ {package_name} - Available")
        except ImportError:
            if package_name in ['python-docx2txt']:
                optional_missing.append(package_name)
                print(f"⚠ {package_name} - Missing (optional, but recommended)")
            else:
                missing.append(package_name)
                print(f"✗ {package_name} - Missing (required)")
    
    return missing, optional_missing

def check_project_structure():
    """Check if project structure is correct"""
    required_files = [
        'src/doc_classify_keywords.py',
        'src/__init__.py',
        'requirements.txt',
        'run.py'
    ]
    
    missing_files = []
    base_dir = Path(__file__).parent
    
    for file_path in required_files:
        full_path = base_dir / file_path
        if full_path.exists():
            print(f"✓ {file_path} - Found")
        else:
            missing_files.append(file_path)
            print(f"✗ {file_path} - Missing")
    
    return missing_files

def test_import():
    """Test if the main module can be imported"""
    try:
        sys.path.insert(0, str(Path(__file__).parent / "src"))
        from src.doc_classify_keywords import get_default_config, read_doc_content
        print("✓ Main module imports successfully")
        
        # Test configuration
        config = get_default_config()
        print("✓ Configuration loaded successfully")
        print(f"  - Categories defined: {len(config['categories'])}")
        print(f"  - Source directory: {config['directories']['source_dir']}")
        
        return True
    except Exception as e:
        print(f"✗ Import failed: {e}")
        return False

def main():
    """Main setup check function"""
    print("=== Simple Document Analysis for Windows - Setup Check ===\n")
    
    all_good = True
    
    # Check Python version
    print("1. Checking Python version...")
    if not check_python_version():
        all_good = False
    print()
    
    # Check dependencies
    print("2. Checking dependencies...")
    missing, optional_missing = check_dependencies()
    if missing:
        all_good = False
        print(f"\nRequired packages missing: {', '.join(missing)}")
        print("Install with: pip install " + " ".join(missing))
    if optional_missing:
        print(f"\nOptional packages missing: {', '.join(optional_missing)}")
        print("Install with: pip install " + " ".join(optional_missing))
    print()
    
    # Check project structure
    print("3. Checking project structure...")
    missing_files = check_project_structure()
    if missing_files:
        all_good = False
        print(f"\nMissing files: {', '.join(missing_files)}")
    print()
    
    # Test imports
    print("4. Testing module imports...")
    if not test_import():
        all_good = False
    print()
    
    # Final result
    print("=" * 60)
    if all_good:
        print("✓ Setup verification PASSED - Ready to run!")
        print("\nTo start processing:")
        print("python run.py")
    else:
        print("✗ Setup verification FAILED - Please fix the issues above")
        print("\nCommon fixes:")
        print("pip install -r requirements.txt")
    print("=" * 60)

if __name__ == "__main__":
    main() 