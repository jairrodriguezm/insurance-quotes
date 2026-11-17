# -*- coding: utf-8 -*-
"""Cloud Storage service — uploads generated documents and returns public URLs.

Supports Google Cloud Storage (preferred when GCS_BUCKET_NAME is set)
and Supabase Storage as fallback.
"""
from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from app.core.config import settings
from app.core.exceptions import StorageUploadError

logger = logging.getLogger(__name__)


def _get_content_type(file_path: str) -> str:
    """Resolve the appropriate Office MIME type based on file extension."""
    ext = Path(file_path).suffix.lower()
    if ext == ".pptx":
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if ext == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if ext == ".pdf":
        return "application/pdf"
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _upload_to_gcs(
    file_path: str,
    bucket_name: str,
    folder: str = "comparativos",
) -> str:
    """Upload a file to Google Cloud Storage and return its public URL."""
    from google.cloud import storage

    client = storage.Client()
    bucket = client.bucket(bucket_name)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = Path(file_path).name
    storage_path = f"{folder}/{timestamp}_{filename}"

    blob = bucket.blob(storage_path)
    blob.upload_from_filename(
        file_path,
        content_type=_get_content_type(file_path),
    )

    public_url = f"https://storage.googleapis.com/{bucket_name}/{storage_path}"
    logger.info("Archivo subido a GCS: %s -> %s", filename, public_url)
    return public_url


def _upload_to_supabase_storage(
    file_path: str,
    bucket: str | None = None,
    folder: str = "comparativos",
) -> str:
    """Upload a file to Supabase Storage and return its public URL."""
    from supabase import create_client

    bucket = bucket or settings.SUPABASE_BUCKET
    client = create_client(settings.SUPABASE_URL, settings.SUPABASE_KEY)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = Path(file_path).name
    storage_path = f"{folder}/{timestamp}_{filename}"

    with open(file_path, "rb") as f:
        file_bytes = f.read()

    client.storage.from_(bucket).upload(
        path=storage_path,
        file=file_bytes,
        file_options={
            "content-type": _get_content_type(file_path),
            "upsert": "true",
        },
    )

    public_url = client.storage.from_(bucket).get_public_url(storage_path)
    logger.info("Archivo subido a Supabase: %s -> %s", filename, public_url)
    return public_url


async def upload_to_storage(
    file_path: str,
    bucket: str | None = None,
    folder: str = "comparativos",
) -> str:
    """Upload a file to the configured cloud storage (GCS or Supabase)."""
    try:
        if settings.GCS_BUCKET_NAME:
            target_bucket = bucket or settings.GCS_BUCKET_NAME
            return _upload_to_gcs(file_path, target_bucket, folder)
        elif settings.SUPABASE_URL and settings.SUPABASE_KEY:
            return _upload_to_supabase_storage(file_path, bucket, folder)
        else:
            raise ValueError("No storage backend configured (GCS_BUCKET_NAME or SUPABASE_URL)")
    except Exception as e:
        raise StorageUploadError(
            message=f"Error al subir archivo a almacenamiento en la nube: {e}",
            detail=str(e),
        ) from e


# Backward compatibility alias
upload_to_supabase = upload_to_storage
