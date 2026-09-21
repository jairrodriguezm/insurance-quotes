# -*- coding: utf-8 -*-
"""Supabase Storage service — uploads generated documents and returns public URLs."""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path

from supabase import create_client, Client

from app.core.config import settings
from app.core.exceptions import StorageUploadError

logger = logging.getLogger(__name__)


def _get_supabase_client() -> Client:
    """Create a Supabase client instance."""
    return create_client(settings.SUPABASE_URL, settings.SUPABASE_KEY)


async def upload_to_supabase(
    file_path: str,
    bucket: str | None = None,
    folder: str = "comparativos",
) -> str:
    """Upload a file to Supabase Storage and return its public URL.

    Args:
        file_path: Local path to the file to upload.
        bucket: Supabase Storage bucket name. Defaults to settings.
        folder: Folder prefix within the bucket.

    Returns:
        Public URL of the uploaded file.

    Raises:
        StorageUploadError: If the upload fails.
    """
    bucket = bucket or settings.SUPABASE_BUCKET
    
    try:
        client = _get_supabase_client()
        
        # Generate a unique storage path
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = Path(file_path).name
        storage_path = f"{folder}/{timestamp}_{filename}"
        
        # Read file bytes
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        
        # Upload to Supabase Storage
        result = client.storage.from_(bucket).upload(
            path=storage_path,
            file=file_bytes,
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "upsert": "true",
            },
        )
        
        # Get public URL
        public_url = client.storage.from_(bucket).get_public_url(storage_path)
        
        logger.info("Archivo subido a Supabase: %s -> %s", filename, public_url)
        return public_url
        
    except Exception as e:
        raise StorageUploadError(
            message=f"Error al subir archivo a Supabase Storage: {e}",
            detail=str(e)
        ) from e
