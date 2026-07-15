"""
Liquidador de Bonificación Judicial — Backend API
FastAPI + Supabase (PostgreSQL)
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, Any
import os, json
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

# ── Supabase client ──────────────────────────────────────────────────────────
def get_db():
    url  = os.getenv("SUPABASE_URL")
    key  = os.getenv("SUPABASE_KEY")
    if not url or not key:
        raise HTTPException(503, "Base de datos no configurada. Configure SUPABASE_URL y SUPABASE_KEY.")
    from supabase import create_client
    return create_client(url, key)

# ── Modelos ──────────────────────────────────────────────────────────────────
class HistorialEntry(BaseModel):
    nip:            str
    nombre:         str
    años:           str
    mes_ejec:       str
    total_bruto:    float
    total_neto:     float
    total_indexado: float
    snap:           str          # JSON completo del estado

class CargoUpdate(BaseModel):
    entidad:      str
    cargo:        str
    año:          int
    bonificacion: float
    salario:      float

class IPCUpdate(BaseModel):
    año:    int
    valores: dict   # {Ene: 140.5, Feb: 141.2, ...}

class TasaUpdate(BaseModel):
    tabla:  str     # 'dtf' o 'corr'
    año:    int
    mes:    int
    valor:  float

# ── HEALTH ───────────────────────────────────────────────────────────────────
@app.get("/api/health")
def health():
    return {"status": "ok", "version": "1.0.0"}

# ── HISTORIAL ────────────────────────────────────────────────────────────────
@app.get("/api/historial")
def get_historial(nip: Optional[str] = None):
    """Obtener historial. Si se pasa nip, filtra por ese NIP."""
    db = get_db()
    q = db.table("historial").select("*").order("created_at", desc=True).limit(100)
    if nip:
        q = q.eq("nip", nip)
    res = q.execute()
    return {"data": res.data}

@app.post("/api/historial")
def save_historial(entry: HistorialEntry):
    """Guardar una liquidación en el historial."""
    db = get_db()
    row = {
        "nip":            entry.nip,
        "nombre":         entry.nombre,
        "años":           entry.años,
        "mes_ejec":       entry.mes_ejec,
        "total_bruto":    entry.total_bruto,
        "total_neto":     entry.total_neto,
        "total_indexado": entry.total_indexado,
        "snap":           entry.snap,
        "created_at":     datetime.utcnow().isoformat()
    }
    res = db.table("historial").insert(row).execute()
    return {"ok": True, "id": res.data[0]["id"] if res.data else None}

@app.delete("/api/historial/{id}")
def delete_historial(id: int):
    db = get_db()
    db.table("historial").delete().eq("id", id).execute()
    return {"ok": True}

@app.delete("/api/historial")
def clear_historial():
    db = get_db()
    db.table("historial").delete().neq("id", 0).execute()
    return {"ok": True}

# ── CARGOS ───────────────────────────────────────────────────────────────────
@app.get("/api/cargos")
def get_cargos():
    """Retorna todos los cargos con bonificación y salario."""
    db = get_db()
    res = db.table("cargos").select("*").order("entidad").order("cargo").order("año").execute()
    return {"data": res.data}

@app.post("/api/cargos")
def upsert_cargo(c: CargoUpdate):
    """Crear o actualizar un cargo/año."""
    db = get_db()
    row = {"entidad": c.entidad, "cargo": c.cargo, "año": c.año,
           "bonificacion": c.bonificacion, "salario": c.salario}
    res = db.table("cargos").upsert(row, on_conflict="entidad,cargo,año").execute()
    return {"ok": True}

@app.delete("/api/cargos")
def delete_cargo(entidad: str, cargo: str, año: int):
    db = get_db()
    db.table("cargos").delete().eq("entidad", entidad).eq("cargo", cargo).eq("año", año).execute()
    return {"ok": True}

# ── IPC ──────────────────────────────────────────────────────────────────────
@app.get("/api/ipc")
def get_ipc():
    db = get_db()
    res = db.table("ipc").select("*").order("año").execute()
    # Convertir a formato {año: {Ene: x, Feb: y, ...}}
    resultado = {}
    for row in res.data:
        resultado[str(row["año"])] = {
            k: row[k] for k in ["Ene","Feb","Mar","Abr","May","Jun",
                                  "Jul","Ago","Sep","Oct","Nov","Dic"] if row.get(k)
        }
    return {"data": resultado}

@app.post("/api/ipc")
def upsert_ipc(entry: IPCUpdate):
    db = get_db()
    row = {"año": entry.año, **entry.valores}
    db.table("ipc").upsert(row, on_conflict="año").execute()
    return {"ok": True}

# ── TASAS ─────────────────────────────────────────────────────────────────────
@app.get("/api/tasas")
def get_tasas():
    db = get_db()
    dtf  = db.table("tasas").select("*").eq("tabla","dtf").execute()
    corr = db.table("tasas").select("*").eq("tabla","corr").execute()
    dtf_dict  = {str(r["año"]*100+r["mes"]): r["valor"] for r in dtf.data}
    corr_dict = {str(r["año"]*100+r["mes"]): r["valor"] for r in corr.data}
    return {"dtf": dtf_dict, "corr": corr_dict}

@app.post("/api/tasas")
def upsert_tasa(t: TasaUpdate):
    db = get_db()
    row = {"tabla": t.tabla, "año": t.año, "mes": t.mes, "valor": t.valor}
    db.table("tasas").upsert(row, on_conflict="tabla,año,mes").execute()
    return {"ok": True}

@app.delete("/api/tasas")
def delete_tasa(tabla: str, año: int, mes: int):
    db = get_db()
    db.table("tasas").delete().eq("tabla", tabla).eq("año", año).eq("mes", mes).execute()
    return {"ok": True}

# ── Servir frontend ──────────────────────────────────────────────────────────
frontend_path = os.path.join(os.path.dirname(__file__), "..", "frontend")
if os.path.exists(frontend_path):
    app.mount("/", StaticFiles(directory=frontend_path, html=True), name="frontend")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT", 8000)), reload=True)
