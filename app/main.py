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

from .services import extract_text_from_pptx, translate_text, create_excel_file, translate_pptx_in_place
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
    slide_range: Optional[str] = Form(None),
    output_format: Optional[str] = Form("pptx"),
    mirror_layout: Optional[str] = Form("true")
):
    """
    Upload a PPTX file, translate it, and return the result.

    Args:
        file: The PPTX file to upload
        slide_range: Optional slide range (e.g., "1-10", "1,3,5", "all")
        output_format: Output format - "pptx" (translated PPTX), "excel" (Excel only), or "both"
        mirror_layout: Whether to mirror layout for RTL (default "true")

    Returns:
        JSON with file_id and filenames for download
    """
    # Validate file type
    if not file.filename.endswith(('.pptx', '.PPTX')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload a .pptx file"
        )

    # Validate output format
    if output_format not in ["pptx", "excel", "both"]:
        output_format = "pptx"

    # Convert mirror_layout string to boolean
    do_mirror = mirror_layout.lower() in ("true", "1", "yes", "on") if mirror_layout else True

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
        result = {
            "success": True,
            "file_id": file_id,
            "output_format": output_format
        }

        # Generate translated PPTX (for "pptx" or "both")
        if output_format in ["pptx", "both"]:
            pptx_output_filename = f"{file_id}_translated.pptx"
            pptx_output_path = OUTPUT_DIR / pptx_output_filename

            translation_result = translate_pptx_in_place(
                str(upload_path),
                str(pptx_output_path),
                slide_range=slide_range,
                mirror_layout=do_mirror
            )

            result["pptx_filename"] = pptx_output_filename
            result["total_phrases"] = translation_result["total_translations"]
            result["total_slides"] = translation_result["total_slides"]
            result["processed_slides"] = translation_result["processed_slides"]

            # If also need Excel, create it from the translation results
            if output_format == "both":
                excel_output_filename = f"{file_id}_translations.xlsx"
                excel_output_path = OUTPUT_DIR / excel_output_filename
                excel_data = [
                    (t["slide"], t["original"], t["translated"])
                    for t in translation_result["translations"]
                ]
                create_excel_file(excel_data, str(excel_output_path))
                result["excel_filename"] = excel_output_filename

            result["message"] = f"Successfully translated {translation_result['total_translations']} phrases in {translation_result['processed_slides']} slides"

        # Generate Excel only (legacy mode)
        elif output_format == "excel":
            extracted_texts = extract_text_from_pptx(str(upload_path), slide_range=slide_range)

            if not extracted_texts:
                raise HTTPException(
                    status_code=400,
                    detail="No translatable text found in the PowerPoint file"
                )

            translations = []
            for slide_num, original_text in extracted_texts:
                translated_text = translate_text(original_text)
                translations.append((slide_num, original_text, translated_text))

            excel_output_filename = f"{file_id}_translations.xlsx"
            excel_output_path = OUTPUT_DIR / excel_output_filename
            create_excel_file(translations, str(excel_output_path))

            result["excel_filename"] = excel_output_filename
            result["total_phrases"] = len(translations)
            result["message"] = f"Successfully processed {len(translations)} phrases"

        # Clean up uploaded file
        os.remove(upload_path)

        return result

    except HTTPException:
        raise
    except Exception as e:
        # Clean up on error
        if upload_path.exists():
            os.remove(upload_path)
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")


@app.get("/api/download/{file_id}")
async def download_file(file_id: str, file_type: Optional[str] = "pptx"):
    """
    Download the generated file (PPTX or Excel).

    Args:
        file_id: The file ID returned from upload
        file_type: Type of file to download - "pptx" or "excel"

    Returns:
        File download (PPTX or Excel)
    """
    if file_type == "excel":
        output_filename = f"{file_id}_translations.xlsx"
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        output_filename = f"{file_id}_translated.pptx"
        media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    output_path = OUTPUT_DIR / output_filename

    if not output_path.exists():
        # Try alternative filename for backward compatibility
        if file_type != "excel":
            alt_filename = f"{file_id}_translations.xlsx"
            alt_path = OUTPUT_DIR / alt_filename
            if alt_path.exists():
                return FileResponse(
                    path=str(alt_path),
                    filename=alt_filename,
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        raise HTTPException(
            status_code=404,
            detail="File not found. It may have expired or been deleted."
        )

    return FileResponse(
        path=str(output_path),
        filename=output_filename,
        media_type=media_type
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
