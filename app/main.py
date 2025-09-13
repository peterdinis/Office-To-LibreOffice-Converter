"""
FastAPI Office to LibreOffice Converter with CORS and Rate Limiting

Converts Microsoft Office files to LibreOffice formats:
- Excel (.xlsx/.xls/.xlsm/.xlsb/.xltx/.xltm) → ODS
- Word (.docx/.doc/.dotx/.dotm) → ODT
- PowerPoint (.pptx/.ppt/.ppsx/.pps/.potx/.potm) → ODP
- Publisher (.pub) → ODT/ODP
- Access (.mdb/.accdb) → ODS via export

Features:
- CORS enabled for cross-origin requests
- Rate limiting (10 requests per minute per IP)
"""

import subprocess
import tempfile
import os
import logging
from io import BytesIO
from collections import OrderedDict
from datetime import datetime, timedelta
from typing import Dict

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# Document processing libraries
from docx import Document
from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.text import P
from odf.draw import Page, Frame, TextBox
import openpyxl
from pyexcel_ods import save_data
from pptx import Presentation

# Configure logging for debugging and monitoring
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Rate limiting storage (in production, use Redis or similar)
rate_limit_storage: Dict[str, list] = {}

# FastAPI application instance
app = FastAPI(
    title="Office to LibreOffice Converter",
    description="""
API allows converting Microsoft Office files to LibreOffice formats:

- Excel (.xlsx/.xls/.xlsm/.xlsb/.xltx/.xltm) → ODS  
- Word (.docx/.doc/.dotx/.dotm) → ODT  
- PowerPoint (.pptx/.ppt/.ppsx/.pps/.potx/.potm) → ODP  
- Publisher (.pub) → ODT/ODP  
- Access (.mdb/.accdb) → ODS via export

Features:
- CORS enabled for cross-origin requests
- Rate limiting: 10 requests per minute per IP address
""",
    version="2.1.0",
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify exact origins
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# Rate limiting configuration
RATE_LIMIT_REQUESTS = 10  # Maximum requests per window
RATE_LIMIT_WINDOW = 60   # Time window in seconds (1 minute)

# Supported file formats configuration
# These formats can be processed directly with Python libraries
PYTHON_SUPPORTED = {
    "excel": ["xlsx", "xls", "xlsm"],
    "word": ["docx"],
    "powerpoint": ["pptx", "ppt", "ppsx", "pps"]
}

# These formats require LibreOffice CLI conversion
LIBRE_SUPPORTED = {
    "excel": ["xlsb", "xltx", "xltm"],
    "word": ["doc", "dotx", "dotm"],
    "powerpoint": ["pps", "ppsx", "potx", "potm"],
    "publisher": ["pub"],
    "access": ["mdb", "accdb"]
}


def check_rate_limit(client_ip: str) -> bool:
    """
    Check if client has exceeded rate limit.
    
    Args:
        client_ip: Client IP address
        
    Returns:
        bool: True if within rate limit, False if exceeded
    """
    now = datetime.now()
    
    # Initialize client record if not exists
    if client_ip not in rate_limit_storage:
        rate_limit_storage[client_ip] = []
    
    # Remove old requests outside the time window
    rate_limit_storage[client_ip] = [
        timestamp for timestamp in rate_limit_storage[client_ip]
        if (now - timestamp).total_seconds() < RATE_LIMIT_WINDOW
    ]
    
    # Check if within rate limit
    if len(rate_limit_storage[client_ip]) >= RATE_LIMIT_REQUESTS:
        return False
    
    # Add current request timestamp
    rate_limit_storage[client_ip].append(now)
    return True


def get_client_ip(request: Request) -> str:
    """
    Extract client IP address from request.
    
    Args:
        request: FastAPI request object
        
    Returns:
        str: Client IP address
    """
    # Check for forwarded IP first (for reverse proxies)
    forwarded_for = request.headers.get("X-Forwarded-For")
    if forwarded_for:
        return forwarded_for.split(",")[0].strip()
    
    # Check for real IP header
    real_ip = request.headers.get("X-Real-IP")
    if real_ip:
        return real_ip
    
    # Fall back to direct client IP
    return request.client.host


@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "message": "Office to LibreOffice Converter API",
        "version": "2.1.0",
        "features": ["CORS enabled", "Rate limiting"],
        "endpoints": {
            "convert": "/convert/",
            "docs": "/docs",
            "openapi": "/openapi.json"
        }
    }


@app.get("/status/")
async def status():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "rate_limit": {
            "requests_per_minute": RATE_LIMIT_REQUESTS,
            "window_seconds": RATE_LIMIT_WINDOW
        }
    }


@app.post("/convert/")
async def convert(request: Request, file: UploadFile = File(..., description="Microsoft Office file to convert")):
    """
    Convert Microsoft Office files to LibreOffice formats.
    
    Args:
        request: FastAPI request object (for rate limiting)
        file: Uploaded Office file
        
    Returns:
        StreamingResponse: Converted LibreOffice file
        
    Raises:
        HTTPException: If conversion fails, file format is unsupported, or rate limit exceeded
    """
    # Rate limiting check
    client_ip = get_client_ip(request)
    if not check_rate_limit(client_ip):
        logger.warning(f"Rate limit exceeded for IP: {client_ip}")
        raise HTTPException(
            status_code=429, 
            detail=f"Rate limit exceeded. Maximum {RATE_LIMIT_REQUESTS} requests per {RATE_LIMIT_WINDOW} seconds.",
            headers={"Retry-After": str(RATE_LIMIT_WINDOW)}
        )
    
    logger.info(f"Processing conversion request from IP: {client_ip}")
    
    # Validate file has extension
    if "." not in file.filename:
        raise HTTPException(status_code=400, detail="File must have an extension")

    # Extract filename and extension
    name, ext = file.filename.rsplit(".", 1)
    ext = ext.lower()

    # Read uploaded file content
    try:
        contents = await file.read()
        if not contents:
            raise HTTPException(status_code=400, detail="Uploaded file is empty")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    # Initialize output stream and filename
    output_stream = BytesIO()
    filename = ""

    try:
        # === PYTHON-BASED CONVERSION ===
        # Faster conversion using Python libraries, no external dependencies
        
        if ext in PYTHON_SUPPORTED["excel"]:
            # Convert Excel files to ODS using openpyxl + pyexcel_ods
            logger.info(f"Converting Excel file {file.filename} using Python libraries")
            
            # Load Excel workbook
            wb = openpyxl.load_workbook(BytesIO(contents))
            sheet = wb.active
            
            # Extract data from active sheet
            data = OrderedDict()
            data["Sheet1"] = [list(row) for row in sheet.iter_rows(values_only=True)]
            
            # Save as ODS format
            save_data(output_stream, data)
            filename = f"{name}.ods"

        elif ext in PYTHON_SUPPORTED["word"]:
            # Convert Word documents to ODT using python-docx + odfpy
            logger.info(f"Converting Word file {file.filename} using Python libraries")
            
            # Load Word document
            doc = Document(BytesIO(contents))
            
            # Create ODT document
            odt = OpenDocumentText()
            
            # Copy paragraphs from Word to ODT
            for para in doc.paragraphs:
                odt.text.addElement(P(text=para.text))
            
            # Save ODT file
            odt.save(output_stream)
            filename = f"{name}.odt"

        elif ext in PYTHON_SUPPORTED["powerpoint"]:
            # Convert PowerPoint presentations to ODP using python-pptx + odfpy
            logger.info(f"Converting PowerPoint file {file.filename} using Python libraries")
            
            try:
                # Load PowerPoint presentation
                prs = Presentation(BytesIO(contents))
                odp = OpenDocumentPresentation()
                
                logger.info(f"Processing PowerPoint with {len(prs.slides)} slides")
                
                # Process each slide
                for slide_idx, slide in enumerate(prs.slides):
                    page = Page()
                    shapes_processed = 0
                    
                    try:
                        # Extract text from all shapes on the slide
                        for shape_idx, shape in enumerate(slide.shapes):
                            try:
                                # Check if shape contains text
                                if hasattr(shape, "has_text_frame") and shape.has_text_frame and shape.text_frame:
                                    # Extract text content safely
                                    text = ""
                                    try:
                                        text = shape.text.strip() if shape.text else ""
                                    except Exception:
                                        # Some shapes might have protected or inaccessible text
                                        continue
                                    
                                    # Add text content to ODP slide if not empty
                                    if text:
                                        frame = Frame()
                                        textbox = TextBox()
                                        textbox.addElement(P(text=text))
                                        frame.addElement(textbox)
                                        page.addElement(frame)
                                        shapes_processed += 1
                                        
                            except Exception as shape_error:
                                # Log shape processing errors but continue
                                logger.warning(f"Skipping shape {shape_idx} on slide {slide_idx}: {shape_error}")
                                continue
                                
                    except Exception as slide_error:
                        # Log slide processing errors but continue
                        logger.warning(f"Error processing shapes on slide {slide_idx}: {slide_error}")
                    
                    # Always add the page to presentation, even if empty
                    odp.presentation.addElement(page)
                    logger.info(f"Processed slide {slide_idx} with {shapes_processed} text shapes")
                
                # Save the ODP file with error handling
                try:
                    odp.save(output_stream)
                    output_stream.seek(0)
                    
                    # Verify output is not empty
                    if output_stream.tell() == 0:
                        raise Exception("Generated ODP file is empty")
                        
                    filename = f"{name}.odp"
                    logger.info(f"Successfully converted PowerPoint to {filename}")
                    
                except Exception as save_error:
                    raise Exception(f"Failed to save ODP file: {save_error}")
                
            except Exception as ppt_error:
                logger.error(f"PowerPoint conversion error: {ppt_error}")
                raise HTTPException(status_code=500, detail=f"PowerPoint conversion failed: {str(ppt_error)}")

        # === LIBREOFFICE CLI CONVERSION ===
        # For complex formats that require LibreOffice's full conversion capabilities
        elif any(ext in v for v in LIBRE_SUPPORTED.values()):
            logger.info(f"Converting {file.filename} using LibreOffice CLI")
            
            tmp_in = None
            tmp_out = None
            
            try:
                # Create temporary input file with proper extension
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp_in:
                    tmp_in.write(contents)
                    tmp_in.flush()

                # Determine target output format based on input file type
                if ext in LIBRE_SUPPORTED.get("excel", []):
                    out_ext = "ods"
                elif ext in LIBRE_SUPPORTED.get("word", []) or ext in LIBRE_SUPPORTED.get("publisher", []):
                    out_ext = "odt"
                elif ext in LIBRE_SUPPORTED.get("powerpoint", []):
                    out_ext = "odp"
                elif ext in LIBRE_SUPPORTED.get("access", []):
                    out_ext = "ods"  # Access databases exported as spreadsheets
                else:
                    out_ext = "odt"  # Default fallback

                # Calculate expected output filename from LibreOffice
                expected_output = os.path.join(
                    os.path.dirname(tmp_in.name), 
                    f"{os.path.splitext(os.path.basename(tmp_in.name))[0]}.{out_ext}"
                )

                # Execute LibreOffice conversion command
                result = subprocess.run([
                    "soffice",                              # LibreOffice command
                    "--headless",                           # Run without GUI
                    "--convert-to", out_ext,                # Target format
                    "--outdir", os.path.dirname(tmp_in.name),  # Output directory
                    tmp_in.name                             # Input file
                ], 
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
                )

                # Check conversion success
                if result.returncode != 0:
                    error_msg = result.stderr.decode() if result.stderr else "Unknown LibreOffice error"
                    raise HTTPException(status_code=500, detail=f"LibreOffice conversion failed: {error_msg}")

                # Verify output file was created
                if not os.path.exists(expected_output):
                    raise HTTPException(status_code=500, detail="LibreOffice conversion did not produce output file")

                # Read converted file into response stream
                with open(expected_output, "rb") as f:
                    output_stream.write(f.read())
                output_stream.seek(0)
                filename = f"{name}.{out_ext}"
                tmp_out = expected_output  # Store for cleanup

            finally:
                # Clean up temporary files
                if tmp_in and os.path.exists(tmp_in.name):
                    try:
                        os.remove(tmp_in.name)
                    except Exception as e:
                        logger.warning(f"Could not remove temp input file: {e}")
                        
                if tmp_out and os.path.exists(tmp_out):
                    try:
                        os.remove(tmp_out)
                    except Exception as e:
                        logger.warning(f"Could not remove temp output file: {e}")

        else:
            # Unsupported file format
            logger.warning(f"Unsupported file format: {ext}")
            return JSONResponse(status_code=400, content={"error": "Unsupported file format"})

    except HTTPException:
        # Re-raise HTTP exceptions without modification
        raise
    except Exception as e:
        # Log and convert general exceptions to HTTP exceptions
        logger.error(f"Conversion failed for {file.filename}: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    # === FINAL VALIDATION AND RESPONSE ===
    
    # Verify we have content to return
    output_stream.seek(0, 2)  # Seek to end to get size
    content_length = output_stream.tell()
    output_stream.seek(0)      # Reset to beginning for streaming
    
    if content_length == 0:
        raise HTTPException(status_code=500, detail="Converted file is empty")

    # Set response headers for file download
    headers = {
        "Content-Disposition": f"attachment; filename={filename}",  # Force download
        "X-Conversion-Status": "success",                           # Custom status header
        "X-Rate-Limit-Remaining": str(RATE_LIMIT_REQUESTS - len(rate_limit_storage.get(client_ip, []))),
        "X-Rate-Limit-Reset": str(int((datetime.now() + timedelta(seconds=RATE_LIMIT_WINDOW)).timestamp()))
    }

    logger.info(f"Successfully converted {file.filename} to {filename} ({content_length} bytes) for IP: {client_ip}")

    # Return converted file as streaming response
    return StreamingResponse(
        output_stream,
        media_type="application/octet-stream",  # Generic binary file type
        headers=headers
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)