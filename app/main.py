import subprocess
import tempfile
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from io import BytesIO
from docx import Document
from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.text import P
from odf.draw import Page, Frame, TextBox
import openpyxl
from pyexcel_ods import save_data
from collections import OrderedDict
from pptx import Presentation
import os

app = FastAPI(
    title="Office to LibreOffice Converter",
    description="""
API allows converting Microsoft Office files to LibreOffice formats:

- Excel (.xlsx/.xls/.xlsm/.xlsb/.xltx/.xltm) → ODS  
- Word (.docx/.doc/.dotx/.dotm) → ODT  
- PowerPoint (.pptx/.ppt/.ppsx/.pps/.potx/.potm) → ODP  
- Publisher (.pub) → ODT/ODP  
- Access (.mdb/.accdb) → ODS via export
""",
    version="2.0.1",
)

# Supported formats
PYTHON_SUPPORTED = {
    "excel": ["xlsx", "xls", "xlsm"],
    "word": ["docx"],
    "powerpoint": ["pptx", "ppt", "ppsx", "pps"]
}

LIBRE_SUPPORTED = {
    "excel": ["xlsb", "xltx", "xltm"],
    "word": ["doc", "dotx", "dotm"],
    "powerpoint": ["pps", "ppsx", "potx", "potm"],
    "publisher": ["pub"],
    "access": ["mdb", "accdb"]
}


@app.post("/convert/")
async def convert(file: UploadFile = File(..., description="Microsoft Office file to convert")):
    if "." not in file.filename:
        raise HTTPException(status_code=400, detail="File must have an extension")

    name, ext = file.filename.rsplit(".", 1)
    ext = ext.lower()

    # Read file content
    try:
        contents = await file.read()
        if not contents:
            raise HTTPException(status_code=400, detail="Uploaded file is empty")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    output_stream = BytesIO()
    filename = ""

    try:
        # --- Python-based conversion ---
        if ext in PYTHON_SUPPORTED["excel"]:
            wb = openpyxl.load_workbook(BytesIO(contents))
            sheet = wb.active
            data = OrderedDict()
            data["Sheet1"] = [list(row) for row in sheet.iter_rows(values_only=True)]
            save_data(output_stream, data)
            filename = f"{name}.ods"

        elif ext in PYTHON_SUPPORTED["word"]:
            doc = Document(BytesIO(contents))
            odt = OpenDocumentText()
            for para in doc.paragraphs:
                odt.text.addElement(P(text=para.text))
            odt.save(output_stream)
            filename = f"{name}.odt"

        elif ext in PYTHON_SUPPORTED["powerpoint"]:
            prs = Presentation(BytesIO(contents))
            odp = OpenDocumentPresentation()
            for slide in prs.slides:
                page = Page()
                try:
                    for shape in slide.shapes:
                        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                            text = shape.text.strip() if shape.text else ""
                            if text:
                                frame = Frame()
                                textbox = TextBox()
                                textbox.addElement(P(text=text))
                                frame.addElement(textbox)
                                page.addElement(frame)
                except Exception as e:
                    # ignoruj shapes, ktoré spôsobujú chybu
                    print(f"Warning: Skipping a problematic shape: {e}")
                odp.presentation.addElement(page)
            odp.save(output_stream)
            filename = f"{name}.odp"

        # --- LibreOffice CLI conversion ---
        elif any(ext in v for v in LIBRE_SUPPORTED.values()):
            tmp_in = None
            tmp_out = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp_in:
                    tmp_in.write(contents)
                    tmp_in.flush()

                # Determine output extension
                if ext in LIBRE_SUPPORTED.get("excel", []):
                    out_ext = "ods"
                elif ext in LIBRE_SUPPORTED.get("word", []) or ext in LIBRE_SUPPORTED.get("publisher", []):
                    out_ext = "odt"
                elif ext in LIBRE_SUPPORTED.get("powerpoint", []):
                    out_ext = "odp"
                elif ext in LIBRE_SUPPORTED.get("access", []):
                    out_ext = "ods"
                else:
                    out_ext = "odt"

                tmp_out = f"{tmp_in.name}_converted.{out_ext}"

                result = subprocess.run(
                    ["soffice", "--headless", "--convert-to", out_ext, "--outdir", os.path.dirname(tmp_in.name), tmp_in.name],
                    check=False,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )

                if result.returncode != 0:
                    raise HTTPException(status_code=500, detail=f"LibreOffice conversion failed: {result.stderr.decode()}")

                with open(tmp_out, "rb") as f:
                    output_stream.write(f.read())
                output_stream.seek(0)
                filename = f"{name}.{out_ext}"

            finally:
                # Cleanup temp files
                if tmp_in and os.path.exists(tmp_in.name):
                    os.remove(tmp_in.name)
                if tmp_out and os.path.exists(tmp_out):
                    os.remove(tmp_out)

        else:
            return JSONResponse(status_code=400, content={"error": "Unsupported file format"})

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    headers = {
        "Content-Disposition": f"attachment; filename={filename}",
        "X-Conversion-Status": "success"
    }

    return StreamingResponse(
        output_stream,
        media_type="application/octet-stream",
        headers=headers
    )
