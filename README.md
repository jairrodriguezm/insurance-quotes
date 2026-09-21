# Comparativo Cotizaciones API

Microservicio FastAPI para la ingesta de cotizaciones de seguros (PDF/DOCX), extracción estructurada con IA (Gemini), y generación automatizada de documentos comparativos Word.

**Multiriesgos de Colombia Ltda.**

## Stack Tecnológico

| Componente | Tecnología |
|---|---|
| Framework | FastAPI + Uvicorn |
| Validación | Pydantic v2 |
| IA / LLM | Google Gemini (`google-genai`) |
| PDF Parsing | pdfplumber |
| DOCX Parsing | python-docx |
| Word Rendering | python-docx (refactored) |
| HTTP Client | httpx |
| Storage | Supabase Storage |

## Arquitectura

```
POST /api/v1/comparatives/jobs
    ↓ (HTTP 202 + job_id)
BackgroundTask:
    ├── 1. Descargar/leer archivos
    ├── 2. Parsear texto (PDF/DOCX)
    ├── 3. MAP: Extraer datos con Gemini (paralelo)
    ├── 4. Extraer metadata del proyecto
    ├── 5. REDUCE: Consolidar + recomendación IA
    ├── 6. Renderizar Word (.docx)
    ├── 7. Subir a Supabase Storage
    └── 8. Callback POST → BPMS/Frontend
```

## Inicio Rápido

### 1. Configurar entorno

```bash
cp .env.example .env
# Editar .env con tus credenciales
```

### 2. Instalar dependencias

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Linux/Mac

pip install -r requirements.txt
```

### 3. Ejecutar

```bash
uvicorn app.main:app --reload --port 8000
```

### 4. Probar

Abrir [http://localhost:8000/docs](http://localhost:8000/docs) para la documentación interactiva (Swagger UI).

#### Enviar cotizaciones por archivo:

```bash
curl -X POST http://localhost:8000/api/v1/comparatives/jobs \
  -F "process_id=test-001" \
  -F "callback_url=https://webhook.site/your-id" \
  -F "files=@CHUBB.pdf" \
  -F "files=@COTIZACION_AXA.pdf" \
  -F "files=@Cotizacion_HDI.docx"
```

#### Enviar cotizaciones por URL:

```bash
curl -X POST http://localhost:8000/api/v1/comparatives/jobs \
  -F "process_id=test-002" \
  -F "callback_url=https://webhook.site/your-id" \
  -F 'file_urls=["https://storage.example.com/chubb.pdf","https://storage.example.com/axa.pdf"]'
```

#### Consultar estado:

```bash
curl http://localhost:8000/api/v1/comparatives/jobs/{job_id}/status
```

## Estructura del Proyecto

```
├── app/
│   ├── api/v1/
│   │   ├── endpoints/
│   │   │   └── comparative.py    # POST /jobs + GET /jobs/{id}/status
│   │   └── api.py                # Router v1
│   ├── core/
│   │   ├── config.py             # Settings (Pydantic BaseSettings)
│   │   └── exceptions.py         # Excepciones de dominio
│   ├── schemas/
│   │   ├── quote.py              # Esquemas para extracción LLM
│   │   └── comparative.py        # Esquemas consolidados + API
│   ├── services/
│   │   ├── document_parser.py    # PDF/DOCX → texto plano
│   │   ├── llm_extractor.py      # Gemini structured outputs
│   │   ├── consolidator.py       # Map-Reduce consolidación
│   │   ├── docx_renderer.py      # Generador Word (6 secciones + anexos)
│   │   ├── storage.py            # Supabase Storage upload
│   │   └── job_runner.py         # Orquestador del pipeline
│   └── main.py                   # FastAPI app
├── assets/                       # Logos y plantilla membrete
├── tests/
├── .env.example
├── Dockerfile
├── requirements.txt
└── README.md
```

## Endpoints

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/api/v1/comparatives/jobs` | Crear job de comparativo |
| `GET` | `/api/v1/comparatives/jobs/{job_id}/status` | Consultar estado |
| `GET` | `/health` | Health check |
| `GET` | `/docs` | Swagger UI |

## Respuesta del Callback

Al completar o fallar, el servicio envía un POST al `callback_url`:

```json
{
  "job_id": "uuid",
  "process_id": "external-id",
  "status": "completed",
  "download_url": "https://supabase.co/storage/v1/..."
}
```

En caso de error:

```json
{
  "job_id": "uuid",
  "process_id": "external-id",
  "status": "failed",
  "error": "Descripción del error"
}
```

## Docker

```bash
docker build -t comparativo-api .
docker run -p 8080:8080 --env-file .env comparativo-api
```

## Variables de Entorno

| Variable | Requerida | Default | Descripción |
|---|---|---|---|
| `GEMINI_API_KEY` | ✅ | — | API key de Google AI Studio |
| `GEMINI_MODEL` | ❌ | `gemini-3.6-flash` | Modelo Gemini a usar |
| `SUPABASE_URL` | ✅ | — | URL del proyecto Supabase |
| `SUPABASE_KEY` | ✅ | — | Anon key de Supabase |
| `SUPABASE_BUCKET` | ❌ | `comparativos` | Bucket de Storage |
| `LOG_LEVEL` | ❌ | `INFO` | Nivel de logging |
| `MAX_FILE_SIZE_MB` | ❌ | `50` | Tamaño máximo por archivo |
