# Simple Document Analysis for Windows

A high-performance Python tool for automatically classifying Word documents (.doc/.docx) based on keyword patterns. Designed specifically for Windows environments with multiprocessing support for fast batch processing.

## Features

- **Fast Processing**: Multiprocessing support for parallel document analysis
- **Multiple Formats**: Supports both .doc and .docx files
- **Smart Classification**: AND logic for precise keyword matching
- **Resume Support**: Automatic progress saving and resume functionality
- **Detailed Logging**: Comprehensive error logging and statistics
- **Windows Optimized**: Uses multiple fallback methods for .doc files

## System Requirements

- **Operating System**: Windows 7/8/10/11
- **Python**: 3.8 or higher
- **Memory**: At least 4GB RAM recommended
- **Storage**: SSD recommended for best performance

## Installation

### 1. Clone or Download the Project

```bash
git clone <repository-url>
cd simple_doc_analysis_win
```

### 2. Create Virtual Environment (Recommended)

```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Optional: Install antiword (for faster .doc processing)

Download antiword from: http://www.winfield.demon.nl/
This can improve .doc file processing speed by 5-10x.

## Configuration

Edit the paths in `src/doc_classify_keywords.py` in the `get_default_config()` function:

```python
'directories': {
    'source_dir': r"C:\path\to\your\documents",      # Input folder
    'dest_base_dir': r"C:\path\to\classified\docs",  # Output folder
    'logs_dir': "logs"
},
```

## Usage

### Quick Start

```bash
python run.py
```

### Performance Testing

```bash
python test_performance.py
```

### Customizing Categories

Modify the `categories` section in `get_default_config()` to define your classification rules:

```python
'YourCategory': {
    'include': ['keyword1', 'keyword2'],  # ALL must be present
    'exclude': ['badword1', 'badword2']   # NONE should be present
},
```

## Performance Features

- **Multiprocessing**: Automatically uses 2-8 CPU cores
- **Smart .doc Reading**: Tries antiword → docx2txt → COM fallback
- **Optimized Search**: Early exit on exclude keyword matches
- **No Artificial Delays**: Removed all sleep() calls
- **Batch Processing**: Reduced I/O frequency

## Expected Performance

- **DOCX files**: 5-15 files/second
- **DOC files**: 2-8 files/second (depends on fallback method)
- **With antiword**: Up to 10x faster for .doc files
- **Multiprocessing**: 3-8x speedup (depends on CPU cores)

## Troubleshooting

### Common Issues

1. **ImportError**: Install requirements with `pip install -r requirements.txt`
2. **Slow .doc processing**: Install antiword or python-docx2txt
3. **Path errors**: Ensure source directory exists and paths use raw strings (r"path")
4. **Memory issues**: Reduce worker count in `get_optimal_worker_count()`

### Performance Tips

1. **Use SSD storage** for input/output directories
2. **Add antivirus exclusions** for document directories
3. **Close other applications** during large batch processing
4. **Use raw strings** for Windows paths: `r"C:\path\to\folder"`

## Project Structure

```
simple_doc_analysis_win/
├── src/
│   ├── __init__.py
│   └── doc_classify_keywords.py    # Main processing module
├── tests/                          # Test directory (empty)
├── run.py                         # Main entry script
├── test_performance.py           # Performance testing
├── requirements.txt              # Dependencies
└── README.md                    # This file
```

## Log Files

The program creates detailed logs in the `logs/` subdirectory:

- `doc_classification_YYYYMMDD_HHMMSS.log` - Error log
- `stats_YYYYMMDD_HHMMSS.json` - Processing statistics
- `processing_progress.pkl` - Resume progress (auto-deleted on completion)

## License

MIT License - See source code for details.

## Support

For Windows-specific issues or performance questions, check the error logs in the `logs/` directory for detailed information. 