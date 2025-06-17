#!/usr/bin/env python3
"""
Performance testing script for document analysis
"""

import time
import os
from pathlib import Path
from src.doc_classify_keywords import read_doc_content, classify_document, get_default_config

def test_single_file_performance(test_file_path: Path, iterations: int = 5):
    """Test processing speed of a single file"""
    print(f"Testing performance with file: {test_file_path}")
    
    if not test_file_path.exists():
        print(f"Test file not found: {test_file_path}")
        return
    
    config = get_default_config()
    categories = config['categories']
    
    total_time = 0
    successful_runs = 0
    
    for i in range(iterations):
        try:
            start_time = time.time()
            
            # Test document reading
            content, error_msg = read_doc_content(test_file_path)
            
            if content and not error_msg:
                # Test classification
                category = classify_document(content, categories)
                
                end_time = time.time()
                processing_time = end_time - start_time
                total_time += processing_time
                successful_runs += 1
                
                print(f"Run {i+1}: {processing_time:.3f}s - Category: {category}")
            else:
                print(f"Run {i+1}: Failed - {error_msg}")
                
        except Exception as e:
            print(f"Run {i+1}: Error - {str(e)}")
    
    if successful_runs > 0:
        avg_time = total_time / successful_runs
        print(f"\nPerformance Summary:")
        print(f"Average processing time: {avg_time:.3f} seconds")
        print(f"Estimated files per minute: {60/avg_time:.1f}")
        print(f"Successful runs: {successful_runs}/{iterations}")
    else:
        print("No successful runs to analyze")

def benchmark_optimization():
    """Benchmark the optimization improvements"""
    print("=== Document Analysis Performance Benchmark ===\n")
    
    # Test with different file types if available
    test_files = [
        Path("test_files/sample.docx"),
        Path("test_files/sample.doc"),
        # Add more test files as needed
    ]
    
    for test_file in test_files:
        if test_file.exists():
            test_single_file_performance(test_file, iterations=3)
            print("-" * 50)
        else:
            print(f"Skipping {test_file} - file not found")
    
    print("\nOptimizations implemented:")
    print("✓ Multiprocessing support (uses multiple CPU cores)")
    print("✓ Faster .doc file reading (tries antiword and docx2txt first)")
    print("✓ Optimized keyword searching (early exit on exclude matches)")
    print("✓ Removed artificial delays")
    print("✓ Reduced progress update frequency")
    print("✓ Optimized DOCX content extraction")

if __name__ == "__main__":
    benchmark_optimization() 