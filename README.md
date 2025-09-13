# Office to LibreOffice Converter

FastAPI application for converting Microsoft Office files to LibreOffice formats.

## Supported Formats

### Excel → ODS
- `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.xltx`, `.xltm`

### Word → ODT  
- `.docx`, `.doc`, `.dotx`, `.dotm`

### PowerPoint → ODP
- `.pptx`, `.ppt`, `.ppsx`, `.pps`, `.potx`, `.potm`

### Publisher → ODT/ODP
- `.pub`

### Access → ODS
- `.mdb`, `.accdb` (exported as spreadsheets)

## Installation & Setup

### Requirements
```bash
pip install fastapi uvicorn python-docx odfpy openpyxl pyexcel-ods python-pptx
```

### Running the Server
```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Server will be available at: `http://localhost:8000`

## API Endpoint

### POST `/convert/`

Converts Microsoft Office files to LibreOffice formats.

**Parameters:**
- `file`: Uploaded Office file (multipart/form-data)

**Response:**
- Converted LibreOffice file for download
- Content-Type: `application/octet-stream`
- Header: `Content-Disposition: attachment; filename=name.ods/odt/odp`

**Example usage with curl:**
```bash
curl -X POST "http://localhost:8000/convert/" \
     -H "accept: application/json" \
     -H "Content-Type: multipart/form-data" \
     -F "file=@document.docx"
```

## Conversion Methods

The application uses two approaches:

### 1. Python Libraries (faster)
- Excel: `openpyxl` + `pyexcel-ods`
- Word: `python-docx` + `odfpy` 
- PowerPoint: `python-pptx` + `odfpy`

### 2. LibreOffice CLI (for complex formats)
- Requires LibreOffice installation
- Used for `.xlsb`, `.doc`, `.pub`, `.mdb` and others

## API Documentation

Interactive documentation available at:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Notes

- PowerPoint conversion extracts text content only
- Complex formats require LibreOffice installed on the system
- Application automatically cleans up temporary files after conversion
- Maximum file size depends on FastAPI configuration

## Logging

The application logs:
- Conversion information
- File processing errors  
- Temporary file warnings

Log level is set to `INFO`.