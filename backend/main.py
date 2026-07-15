"""
Liquidador de Bonificación Judicial — Backend API
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
from typing import Optional
import os
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="Liquidador Bonificación Judicial", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_db():
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    if not url or not key:
        raise HTTPException(503, "Base de datos no configurada.")
    from supabase import create_client
    return create_client(url, key)

class HistorialEntry(BaseModel):
    nip: str
    nombre: str
    años: str
    mes_ejec: str
    total_bruto: float
    total_neto: float
    total_indexado: float
    snap: str

class CargoUpdate(BaseModel):
    entidad: str
    cargo: str
    año: int
    bonificacion: float
    salario: float

class IPCUpdate(BaseModel):
    año: int
    valores: dict

class TasaUpdate(BaseModel):
    tabla: str
    año: int
    mes: int
    valor: float

@app.get("/api/health")
def health():
    return {"status": "ok"}

@app.get("/api/historial")
def get_historial(nip: Optional[str] = None):
    db = get_db()
    q = db.table("historial").select("*").order("created_at", desc=True).limit(100)
    if nip:
        q = q.eq("nip", nip)
    return {"data": q.execute().data}

@app.post("/api/historial")
def save_historial(entry: HistorialEntry):
    db = get_db()
    row = {**entry.dict(), "created_at": datetime.utcnow().isoformat()}
    res = db.table("historial").insert(row).execute()
    return {"ok": True, "id": res.data[0]["id"] if res.data else None}

@app.delete("/api/historial/{id}")
def delete_historial(id: int):
    get_db().table("historial").delete().eq("id", id).execute()
    return {"ok": True}

@app.delete("/api/historial")
def clear_historial():
    get_db().table("historial").delete().neq("id", 0).execute()
    return {"ok": True}

@app.get("/api/cargos")
def get_cargos():
    res = get_db().table("cargos").select("*").order("entidad").order("cargo").order("año").execute()
    return {"data": res.data}

@app.post("/api/cargos")
def upsert_cargo(c: CargoUpdate):
    get_db().table("cargos").upsert(
        {"entidad": c.entidad, "cargo": c.cargo, "año": c.año,
         "bonificacion": c.bonificacion, "salario": c.salario},
        on_conflict="entidad,cargo,año"
    ).execute()
    return {"ok": True}

@app.get("/api/ipc")
def get_ipc():
    res = get_db().table("ipc").select("*").order("año").execute()
    resultado = {}
    for row in res.data:
        resultado[str(row["año"])] = {
            k: row[k] for k in ["Ene","Feb","Mar","Abr","May","Jun",
                                 "Jul","Ago","Sep","Oct","Nov","Dic"] if row.get(k)
        }
    return {"data": resultado}

@app.post("/api/ipc")
def upsert_ipc(entry: IPCUpdate):
    get_db().table("ipc").upsert({"año": entry.año, **entry.valores}, on_conflict="año").execute()
    return {"ok": True}

@app.get("/api/tasas")
def get_tasas():
    db = get_db()
    dtf  = {str(r["año"]*100+r["mes"]): r["valor"] for r in db.table("tasas").select("*").eq("tabla","dtf").execute().data}
    corr = {str(r["año"]*100+r["mes"]): r["valor"] for r in db.table("tasas").select("*").eq("tabla","corr").execute().data}
    return {"dtf": dtf, "corr": corr}

@app.post("/api/tasas")
def upsert_tasa(t: TasaUpdate):
    get_db().table("tasas").upsert(
        {"tabla": t.tabla, "año": t.año, "mes": t.mes, "valor": t.valor},
        on_conflict="tabla,año,mes"
    ).execute()
    return {"ok": True}

@app.delete("/api/tasas")
def delete_tasa(tabla: str, año: int, mes: int):
    get_db().table("tasas").delete().eq("tabla", tabla).eq("año", año).eq("mes", mes).execute()
    return {"ok": True}

# ── Servir frontend ──────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CWD = os.getcwd()

# Buscar index.html en todas las rutas posibles
INDEX_PATHS = [
    os.path.join(CWD, "frontend", "index.html"),
    os.path.join(BASE_DIR, "..", "frontend", "index.html"),
    os.path.join(CWD, "index.html"),
    os.path.join(BASE_DIR, "index.html"),
    "/opt/render/project/src/frontend/index.html",
]

INDEX_FILE = None
for p in INDEX_PATHS:
    if os.path.exists(p):
        INDEX_FILE = os.path.abspath(p)
        print(f"index.html encontrado en: {INDEX_FILE}")
        break

if INDEX_FILE:
    FRONTEND_DIR = os.path.dirname(INDEX_FILE)
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")

    @app.get("/", response_class=HTMLResponse)
    def serve_index():
        with open(INDEX_FILE, encoding="utf-8") as f:
            return f.read()
else:
    print(f"ERROR: index.html no encontrado. CWD={CWD}, BASE={BASE_DIR}")
    print(f"Contenido CWD: {os.listdir(CWD)}")

    @app.get("/")
    def root():
        dirs = {}
        for p in [CWD, BASE_DIR]:
            try:
                dirs[p] = os.listdir(p)
            except:
                dirs[p] = "error"
        return {"error": "index.html no encontrado", "dirs": dirs,
                "index_paths_checked": INDEX_PATHS}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT", 8000)), reload=True)
