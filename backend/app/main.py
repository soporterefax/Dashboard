import os

from dotenv import load_dotenv

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routes.equipos import router as equipos_router
from app.services.excel_service import obtener_estado_fuente


# =========================
# VARIABLES DE ENTORNO
# =========================

load_dotenv()


# =========================
# FASTAPI
# =========================

app = FastAPI(
    title="API Inventario TI",
    version="1.0.0",
)


# =========================
# CORS
# =========================

cors_origins = [
    origin.strip()
    for origin in os.getenv(
        "CORS_ORIGINS",
        "http://localhost:5173"
    ).split(",")
    if origin.strip()
]


app.add_middleware(
    CORSMiddleware,

    allow_origins=cors_origins,

    allow_credentials=False,

    allow_methods=[
        "GET",
        "POST",
        "PUT",
        "PATCH",
        "DELETE",
        "OPTIONS",
    ],

    allow_headers=["*"],
)


# =========================
# ROUTERS
# =========================

app.include_router(
    equipos_router
)


# =========================
# HOME
# =========================

@app.get("/")
def home():

    return {
        "mensaje":
            "API Inventario TI funcionando"
    }


# =========================
# HEALTH
# =========================

@app.get("/health")
def health():

    return {
        "status": "ok",

        "data_source":
            os.getenv(
                "DATA_SOURCE",
                "local"
            ),
    }


# =========================
# FUENTE DE DATOS
# =========================

@app.get("/fuente-datos")
def fuente_datos():

    return obtener_estado_fuente()