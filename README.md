# Ropita · asistente de estilo

Chatbot que recomienda outfits según género, edad, ocasión, clima, día/noche y día de la semana.

👉 **Pruébalo en línea:** https://javier-galindobernal.github.io/ropita/
(en línea funcionan las preguntas y recomendaciones; el chat libre con IA requiere ejecutarlo en tu computador con Ollama)

## Estructura
```
ropita/
├── ropita_db.xlsx        # Base de conocimiento (editable en Excel)
├── generar_excel.py     # Regenera el Excel desde cero
├── server.py            # Backend: exporta el Excel a JSON, sirve docs/ y habla con Ollama
└── docs/                 # Frontend (publicado en GitHub Pages)
    ├── index.html        # Pantalla del chat
    ├── ropita_data.json  # Datos exportados del Excel (se regenera al iniciar server.py)
    ├── ropita-avatar.svg  # Imagen del bot (vectorial)
    └── ropita-avatar-512.png
```

## Cómo ejecutarlo
```bash
pip install -r requirements.txt
ollama pull llama3.2        # si aún no tienes el modelo
python3 server.py           # abre http://localhost:8000
```
Otro modelo: `OLLAMA_MODEL=qwen2.5 python3 server.py`

## Cómo funciona (modo híbrido)
1. **Preguntas predeterminadas** (hoja `Preguntas`): el bot pregunta con botones. Solo muestra opciones que
   existen en la hoja `Outfits` (p. ej. a un niño no le ofrece "Trabajo").
2. **Recomendación exacta** desde la hoja `Outfits` (sin IA, rápida y controlada).
3. **Chat libre con Ollama**: tras la recomendación, el usuario escribe preguntas; el servidor envía a Ollama
   el perfil + el outfit del Excel como contexto, para que la IA no invente fuera de la base.

## Hojas del Excel
| Hoja | Contenido |
|---|---|
| Outfits | 1.856 combinaciones: género × edad × ocasión × clima × día/noche × día de la semana |
| Preguntas | Orden, clave, texto y opciones (separadas por `\|`) |
| Prendas | Catálogo de prendas con formalidad, climas y ocasiones |
| Climas / Ocasiones / Grupos_edad | Tablas de referencia |

Puedes editar filas directamente en Excel y reiniciar `server.py` (regenera `docs/ropita_data.json`).
Para que el cambio se vea en la versión en línea, haz commit y push de `ropita_db.xlsx` y `docs/ropita_data.json`. Si agregas una opción nueva en `Preguntas`,
debe existir en la columna correspondiente de `Outfits`.

## API
- `POST /api/chat` `{respuestas, outfit, mensaje, historial}` → respuesta de Ollama (solo en local)

## Créditos
Creado por **Laura Galindo, Valentina Rodríguez, Sara Cárdenas y Joseth Sierra** — Grado 1002.
