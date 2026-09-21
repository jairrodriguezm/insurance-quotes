# -*- coding: utf-8 -*-
"""FastAPI application entry point — Comparativo de Cotizaciones de Seguros.

A microservice for ingesting insurance quotes (PDF/DOCX), extracting structured
data via AI (Gemini), and generating consolidated comparison Word documents.

Start with:
    uvicorn app.main:app --reload --port 8000
"""
from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.v1.api import api_router
from app.core.config import settings

# --------------------------------------------------------------------------- #
# Logging                                                                      #
# --------------------------------------------------------------------------- #

logging.basicConfig(
    level=getattr(logging, settings.LOG_LEVEL.upper(), logging.INFO),
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------- #
# Application                                                                  #
# --------------------------------------------------------------------------- #

@asynccontextmanager
async def lifespan(app: FastAPI):
    """Lifecycle events for the FastAPI application."""
    # Fail fast if required environment variables are missing
    try:
        settings.validate_required()
        logger.info("Configuration validated successfully.")
    except ValueError as e:
        logger.error(f"Configuration error: {e}")
        # We don't exit here so the health check endpoint can still run,
        # but the jobs will fail if they try to use the APIs.
    
    yield
    # Cleanup on shutdown

app = FastAPI(
    title="Comparativo Cotizaciones API",
    lifespan=lifespan,
    description=(
        "Microservicio para la ingesta de cotizaciones de seguros (PDF/DOCX), "
        "extracción estructurada con IA (Gemini), y generación automatizada de "
        "documentos comparativos Word para Multiriesgos de Colombia Ltda."
    ),
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS — allow all origins for development (restrict in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount API v1 routes
app.include_router(api_router, prefix="/api/v1")


# --------------------------------------------------------------------------- #
# Health check                                                                 #
# --------------------------------------------------------------------------- #

@app.get("/health", tags=["Health"])
async def health_check() -> dict:
    """Health check endpoint for load balancers and monitoring."""
    return {
        "status": "healthy",
        "service": "comparativo-cotizaciones-api",
        "version": "1.0.0",
    }


@app.get("/", tags=["Health"])
async def root() -> dict:
    """Root endpoint with API information."""
    return {
        "service": "Comparativo Cotizaciones API",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/health",
    }
