# -*- coding: utf-8 -*-
"""API v1 router — aggregates all v1 endpoint routers."""
from fastapi import APIRouter

from app.api.v1.endpoints.comparative import router as comparative_router

api_router = APIRouter()
api_router.include_router(comparative_router)
