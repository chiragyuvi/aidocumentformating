from fastapi import HTTPException

from app.core.config import settings


DOCX_SIGNATURE = b"PK"


def validate_docx_upload(filename: str | None, content: bytes) -> None:
    if not filename or not filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    max_size_bytes = settings.max_upload_size_mb * 1024 * 1024
    if len(content) > max_size_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"File '{filename}' exceeds the {settings.max_upload_size_mb} MB limit.",
        )

    if not content.startswith(DOCX_SIGNATURE):
        raise HTTPException(
            status_code=400,
            detail=f"File '{filename}' is not a valid DOCX package.",
        )
