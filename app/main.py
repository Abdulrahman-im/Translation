"""FastAPI application for PPTX translation."""

import os
import uuid
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, Body, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional

from .services import extract_text_from_pptx, translate_text, create_excel_file
from .services.dictionary import get_all_entries, add_entry, add_entries_bulk, get_dictionary_stats


class DictionaryEntry(BaseModel):
    english: str
    arabic: str


class BulkEntriesRequest(BaseModel):
    entries: List[DictionaryEntry]


from .services.alignment import build_dictionary_from_parallel_pptx

# Get base directory
BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

# Ensure directories exist
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(
    title="PPTX Translation Service",
    description="Upload PowerPoint files and get translations in Excel format",
    version="1.0.0"
)

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/upload")
async def upload_pptx(
    file: UploadFile = File(...),
    slide_range: Optional[str] = Form(None)
):
    """
    Upload a PPTX file, extract text, translate, and return file ID.

    Args:
        file: The PPTX file to upload
        slide_range: Optional slide range (e.g., "1-10", "1,3,5", "all")

    Returns:
        JSON with file_id and filename for download
    """
    # Validate file type
    if not file.filename.endswith(('.pptx', '.PPTX')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload a .pptx file"
        )

    # Generate unique file ID
    file_id = str(uuid.uuid4())

    # Save uploaded file
    upload_path = UPLOAD_DIR / f"{file_id}.pptx"
    try:
        content = await file.read()
        with open(upload_path, "wb") as f:
            f.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

    try:
        # Extract text from PPTX (with optional slide range filtering)
        extracted_texts = extract_text_from_pptx(str(upload_path), slide_range=slide_range)

        if not extracted_texts:
            raise HTTPException(
                status_code=400,
                detail="No translatable text found in the PowerPoint file"
            )

        # Translate all texts
        translations = []
        for slide_num, original_text in extracted_texts:
            translated_text = translate_text(original_text)
            translations.append((slide_num, original_text, translated_text))

        # Generate Excel file
        output_filename = f"{file_id}_translations.xlsx"
        output_path = OUTPUT_DIR / output_filename
        create_excel_file(translations, str(output_path))

        # Clean up uploaded file
        os.remove(upload_path)

        return {
            "success": True,
            "file_id": file_id,
            "filename": output_filename,
            "total_phrases": len(translations),
            "message": f"Successfully processed {len(translations)} phrases"
        }

    except HTTPException:
        raise
    except Exception as e:
        # Clean up on error
        if upload_path.exists():
            os.remove(upload_path)
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")


@app.get("/api/download/{file_id}")
async def download_excel(file_id: str):
    """
    Download the generated Excel file.

    Args:
        file_id: The file ID returned from upload

    Returns:
        Excel file download
    """
    output_filename = f"{file_id}_translations.xlsx"
    output_path = OUTPUT_DIR / output_filename

    if not output_path.exists():
        raise HTTPException(
            status_code=404,
            detail="File not found. It may have expired or been deleted."
        )

    return FileResponse(
        path=str(output_path),
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/api/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy"}


# =============================================================================
# Dictionary Management Endpoints
# =============================================================================

@app.get("/api/dictionary")
async def get_dictionary():
    """Get all dictionary entries."""
    entries = get_all_entries()
    stats = get_dictionary_stats()
    return {
        "entries": entries,
        "stats": stats
    }


@app.post("/api/dictionary/add")
async def add_dictionary_entry(entry: DictionaryEntry):
    """Add a single entry to the dictionary (word or sentence)."""
    if not entry.english.strip() or not entry.arabic.strip():
        raise HTTPException(status_code=400, detail="Both English and Arabic text are required")

    success = add_entry(entry.english.strip(), entry.arabic.strip(), validated=True)
    return {"success": success, "message": "Entry added successfully"}


@app.post("/api/dictionary/add-bulk")
async def add_dictionary_entries_bulk(request: BulkEntriesRequest):
    """Add multiple entries to the dictionary at once."""
    if not request.entries:
        raise HTTPException(status_code=400, detail="No entries provided")

    # Filter out empty entries
    valid_entries = [
        {"english": e.english.strip(), "arabic": e.arabic.strip(), "validated": True}
        for e in request.entries
        if e.english.strip() and e.arabic.strip()
    ]

    if not valid_entries:
        raise HTTPException(status_code=400, detail="No valid entries found")

    added_count = add_entries_bulk(valid_entries)

    return {
        "success": True,
        "added": added_count,
        "total_submitted": len(request.entries),
        "message": f"Added {added_count} entries to dictionary"
    }


@app.post("/api/dictionary/build")
async def build_dictionary(
    english_file: UploadFile = File(...),
    arabic_file: UploadFile = File(...)
):
    """
    Build dictionary from parallel PowerPoint files.

    Upload two PPTX files (English and Arabic counterparts) to automatically
    extract and validate translation pairs.
    """
    # Validate file types
    if not english_file.filename.endswith(('.pptx', '.PPTX')):
        raise HTTPException(status_code=400, detail="English file must be a .pptx file")
    if not arabic_file.filename.endswith(('.pptx', '.PPTX')):
        raise HTTPException(status_code=400, detail="Arabic file must be a .pptx file")

    # Generate unique IDs for temp files
    file_id = str(uuid.uuid4())
    english_path = UPLOAD_DIR / f"{file_id}_en.pptx"
    arabic_path = UPLOAD_DIR / f"{file_id}_ar.pptx"

    try:
        # Save uploaded files
        en_content = await english_file.read()
        ar_content = await arabic_file.read()

        with open(english_path, "wb") as f:
            f.write(en_content)
        with open(arabic_path, "wb") as f:
            f.write(ar_content)

        # Build dictionary from parallel files
        result = build_dictionary_from_parallel_pptx(
            str(english_path),
            str(arabic_path),
            validate=True
        )

        return {
            "success": True,
            "total_candidates": result["total_candidates"],
            "validated_pairs": result["validated_pairs"],
            "added_to_dictionary": result["added_to_dictionary"],
            "candidates": result["candidates"],
            "message": f"Processed {result['total_candidates']} pairs, validated {result['validated_pairs']}, added {result['added_to_dictionary']} to dictionary"
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to build dictionary: {str(e)}")

    finally:
        # Clean up temp files
        if english_path.exists():
            os.remove(english_path)
        if arabic_path.exists():
            os.remove(arabic_path)


@app.get("/api/dictionary/stats")
async def dictionary_stats():
    """Get dictionary statistics."""
    return get_dictionary_stats()


# Mount static files last (so API routes take precedence)
app.mount("/", StaticFiles(directory=str(BASE_DIR / "app" / "static"), html=True), name="static")
