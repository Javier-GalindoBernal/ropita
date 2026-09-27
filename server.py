"""Servidor de Ropita: lee ropita_db.xlsx, sirve el frontend y conecta con Ollama.

Uso:  python3 server.py      →  http://localhost:8000
Variables opcionales: OLLAMA_URL (default http://localhost:11434), OLLAMA_MODEL (default llama3.2), PORT.
"""
import json
import os
import re
import urllib.error
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE_DIR, "ropita_db.xlsx")
OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2")
PORT = int(os.environ.get("PORT", 8000))


def leer_hoja(wb, nombre):
    filas = list(wb[nombre].iter_rows(values_only=True))
    return [dict(zip(filas[0], f)) for f in filas[1:] if any(f)]


wb = load_workbook(EXCEL, read_only=True)
OUTFITS = leer_hoja(wb, "Outfits")
PREGUNTAS = sorted(leer_hoja(wb, "Preguntas"), key=lambda p: p["orden"])
PRENDAS = leer_hoja(wb, "Prendas")
print(f"Base cargada: {len(OUTFITS)} outfits, {len(PREGUNTAS)} preguntas, {len(PRENDAS)} prendas")


def limpiar(valor):
    """'Infantil (3-12 años)' -> 'Infantil'"""
    return re.sub(r"\s*\(.*\)$", "", str(valor)).strip()


def filtrar(respuestas):
    claves = {k: limpiar(v) for k, v in respuestas.items() if v}
    return [o for o in OUTFITS if all(o.get(k) == v for k, v in claves.items())]


def opciones_validas(pregunta, respuestas):
    """Solo ofrece opciones que tengan al menos un outfit con las respuestas previas."""
    posibles = {o[pregunta["clave"]] for o in filtrar(respuestas)}
    return [op for op in pregunta["opciones"].split("|") if limpiar(op) in posibles]


SISTEMA = """Eres Ropita, un asistente de moda amable que habla español.
Recomiendas outfits según el clima, la ocasión, el momento del día, el día de la semana, la edad y el género.
Basa tus respuestas en el PERFIL y el OUTFIT BASE que vienen de la base de datos; puedes complementarlos con
alternativas, colores y consejos. Responde en máximo 120 palabras, con viñetas cuando ayuden.
Si la persona es menor de edad, mantén recomendaciones apropiadas para su edad.
Si te preguntan algo que no sea moda o vestuario, redirige amablemente la conversación."""


def chat_ollama(perfil, outfit, mensaje, historial):
    contexto = f"PERFIL: {json.dumps(perfil, ensure_ascii=False)}\nOUTFIT BASE: {json.dumps(outfit, ensure_ascii=False)}"
    mensajes = [{"role": "system", "content": SISTEMA + "\n\n" + contexto}]
    mensajes += historial[-8:]
    mensajes.append({"role": "user", "content": mensaje})
    cuerpo = json.dumps({"model": OLLAMA_MODEL, "messages": mensajes, "stream": False}).encode()
    req = urllib.request.Request(f"{OLLAMA_URL}/api/chat", data=cuerpo, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.load(r)["message"]["content"]


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=os.path.join(BASE_DIR, "static"), **kw)

    def responder(self, datos, status=200):
        cuerpo = json.dumps(datos, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def do_POST(self):
        datos = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or "{}")
        respuestas = datos.get("respuestas", {})

        if self.path == "/api/siguiente":
            # Devuelve la próxima pregunta pendiente con sus opciones válidas
            for p in PREGUNTAS:
                if p["clave"] not in respuestas:
                    return self.responder({"clave": p["clave"], "pregunta": p["pregunta"],
                                           "opciones": opciones_validas(p, respuestas)})
            return self.responder({"fin": True})

        if self.path == "/api/recomendar":
            encontrados = filtrar(respuestas)
            if not encontrados:
                return self.responder({"error": "No encontré un outfit para esa combinación."}, 404)
            return self.responder({"outfit": encontrados[0]})

        if self.path == "/api/chat":
            try:
                texto = chat_ollama(respuestas, datos.get("outfit"), datos.get("mensaje", ""), datos.get("historial", []))
                return self.responder({"respuesta": texto})
            except (urllib.error.URLError, TimeoutError, KeyError) as e:
                return self.responder({"error": f"No pude conectar con Ollama ({OLLAMA_URL}, modelo {OLLAMA_MODEL}): {e}"}, 503)

        self.responder({"error": "Ruta no encontrada"}, 404)


if __name__ == "__main__":
    print(f"Ropita listo en http://localhost:{PORT}  (Ollama: {OLLAMA_URL}, modelo {OLLAMA_MODEL})")
    ThreadingHTTPServer(("", PORT), Handler).serve_forever()
