# DocuX - Document Translation Tool

DocuX is a powerful document translation tool that supports PDF and DOCX files using AWS Translate services. It provides a user-friendly web interface built with Streamlit.

## Features

- PDF to DOCX conversion using Microsoft Word (with LibreOffice fallback)
- Document translation using AWS Translate
- Support for multiple languages
- Clean and intuitive web interface
- Preserves document formatting
- Progress tracking for translation tasks

## Supported Languages

- English
- French
- Spanish
- German
- Italian
- Arabic
- Chinese (Simplified)
- Japanese

## Prerequisites

- Python 3.7+
- Microsoft Word (recommended for PDF conversion)
- LibreOffice (fallback for PDF conversion)
- AWS Account with Translate service access

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd translator
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Set up AWS credentials as environment variables:
```bash
# Windows
set aws_key_id=your_aws_access_key_id
set aws_secret_key=your_aws_secret_access_key

# Linux/MacOS
export aws_key_id=your_aws_access_key_id
export aws_secret_key=your_aws_secret_access_key
```

## Usage

1. Start the Streamlit application:
```bash
streamlit run translator.py
```

2. Access the web interface at `http://localhost:8501`

3. Select source and target languages

4. Upload your PDF or DOCX file

5. Click "Translate" and wait for the process to complete

6. Download your translated document

## Technical Details

- Uses AWS Translate for high-quality translations
- Implements text chunking to handle AWS Translate's character limits
- Preserves document formatting during translation
- Handles both PDF and DOCX file formats
- Uses Microsoft Word COM automation for PDF conversion (when available)

## Dependencies

- `streamlit`: Web interface
- `boto3`: AWS SDK for Python
- `python-docx`: DOCX file handling
- `nltk`: Text processing
- `pywin32`: Microsoft Office automation

## Error Handling

The application includes comprehensive error handling for:
- Missing AWS credentials
- File conversion issues
- Translation errors
- Invalid file formats

## Notes

- PDF conversion works best with Microsoft Word installed
- Large documents may take longer to process
- AWS credentials must have appropriate permissions for Translate service
