"""Custom exceptions for the comparative quote service."""

from __future__ import annotations


class ComparativeServiceError(Exception):
    """Base exception for the comparative service."""

    def __init__(self, message: str, detail: str | None = None) -> None:
        """Initialize the comparative service error.

        Args:
            message: Human-readable error description.
            detail: Optional additional error details or stack trace context.
        """
        self.message = message
        self.detail = detail
        super().__init__(message)


class DocumentParsingError(ComparativeServiceError):
    """Raised when a document cannot be parsed."""

    pass


class LLMExtractionError(ComparativeServiceError):
    """Raised when the LLM fails to extract structured data."""

    pass


class StorageUploadError(ComparativeServiceError):
    """Raised when file upload to Supabase fails."""

    pass


class CallbackError(ComparativeServiceError):
    """Raised when the webhook callback to BPMS fails."""

    pass


class FileDownloadError(ComparativeServiceError):
    """Raised when a remote file cannot be downloaded."""

    pass
