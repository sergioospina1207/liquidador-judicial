"""
Extractor de cargos y fechas de certificados judiciales usando Claude API
"""
import anthropic
import base64
import json
import os
import re
from datetime import datetime

# Lista de cargos de la app para comparar
CARGOS_RAMA = [
    "Abogado Asistente Grado 1","Abogado Asistente Grado 2","Abogado Asistente Grado 3",
    "Asistente Judicial I","Asistente Judicial II","Asistente Judicial III",
    "Asistente Judicial IV","Auxiliar Judicial I","Auxiliar Judicial II",
    "Auxiliar Judicial III","Auxiliar Judicial IV","Citador",
    "Juez Civil del Circuito","Juez Civil Municipal","Juez de Familia",
    "Juez Laboral del Circuito","Juez Municipal","Juez Penal del Circuito",
    "Juez Penal Municipal","Juez Promiscuo Municipal","Juez Promiscuo del Circuito",
    "Juez del Circuito","Magistrado Tribunal","Oficial Mayor",
    "Oficial Mayor del Circuito","Oficial Mayor Municipal",
    "Secretario","Secretario del Circuito","Secretario Municipal",
    "Secretario Seccional","Escribiente","Escribiente del Circuito",
    "Escribiente Municipal","Técnico Judicial I","Técnico Judicial II",
    "Técnico Judicial III","Técnico Judicial IV",
]

CARGOS_FISC = [
    "Asistente de Fiscal","Fiscal Delegado","Fiscal Local",
    "Investigador","Técnico de Fiscalía","Auxiliar de Fiscalía",
    "Secretario de Fiscalía","Profesional Universitario",
]

def comparar_cargo(cargo_cert: str, lista_cargos: list) -> tuple:
    """Compara un cargo del certificado con la lista y devuelve (mejor_match, score)"""
    cargo_norm = cargo_cert.upper().strip()
    # Quitar sufijos numéricos como "00", "01", etc.
    cargo_norm = re.sub(r'\s+\d{2}$', '', cargo_norm).strip()
    
    mejor = None
    mejor_score = 0
    
    for c in lista_cargos:
        c_norm = c.upper()
        # Palabras en común
        words_cert = set(cargo_norm.split())
        words_lista = set(c_norm.split())
        # Quitar artículos
        stopwords = {'DE', 'DEL', 'LA', 'EL', 'LAS', 'LOS', 'Y', 'EN'}
        words_cert -= stopwords
        words_lista -= stopwords
        
        if not words_cert or not words_lista:
            continue
            
        interseccion = words_cert & words_lista
        union = words_cert | words_lista
        score = len(interseccion) / len(union)
        
        if score > mejor_score:
            mejor_score = score
            mejor = c
    
    return mejor, mejor_score

def extraer_cargos_pdf(pdf_bytes: bytes, entidad: str = "RAMA") -> dict:
    """
    Extrae cargos y fechas de un certificado PDF usando Claude.
    Retorna dict con lista de cargos procesados.
    """
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    
    prompt = """Analiza este certificado judicial colombiano y extrae ÚNICAMENTE:
1. Los cargos ejercidos (primera columna)
2. La fecha de inicio de cada cargo
3. La fecha de fin de cada cargo

Responde SOLO con un JSON válido con esta estructura exacta, sin texto adicional:
{
  "cargos": [
    {
      "cargo": "nombre exacto del cargo como aparece en el documento",
      "fecha_inicio": "DD/MM/YYYY",
      "fecha_fin": "DD/MM/YYYY"
    }
  ]
}

Si una fecha no está clara, usa null. Extrae TODOS los registros que aparezcan."""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64,
                    },
                },
                {
                    "type": "text",
                    "text": prompt
                }
            ],
        }]
    )
    
    raw = response.content[0].text.strip()
    # Limpiar posibles backticks
    raw = re.sub(r'^```json?\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    
    data = json.loads(raw)
    cargos_raw = data.get("cargos", [])
    
    # Comparar con lista de cargos de la app
    lista = CARGOS_RAMA if entidad.upper() == "RAMA" else CARGOS_FISC
    resultado = []
    
    for item in cargos_raw:
        cargo_cert = item.get("cargo", "")
        fecha_ini  = item.get("fecha_inicio")
        fecha_fin  = item.get("fecha_fin")
        
        mejor_match, score = comparar_cargo(cargo_cert, lista)
        
        if score >= 0.6:
            estado = "ok"
            cargo_app = mejor_match
        elif score >= 0.3:
            estado = "revisar"
            cargo_app = mejor_match
        else:
            estado = "no_encontrado"
            cargo_app = None
        
        resultado.append({
            "cargo_certificado": cargo_cert,
            "cargo_app":         cargo_app,
            "fecha_inicio":      fecha_ini,
            "fecha_fin":         fecha_fin,
            "estado":            estado,
            "score":             round(score, 2)
        })
    
    return {"cargos": resultado, "total": len(resultado)}
