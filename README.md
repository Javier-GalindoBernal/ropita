# Ropita · asistente de estilo

Chatbot que recomienda outfits según género, edad, ocasión, clima, día/noche y día de la semana.

## Estructura
```
ropita/
├── ropita_db.xlsx        # Base de conocimiento (editable en Excel)
├── generar_excel.py     # Regenera el Excel desde cero
├── server.py            # Backend: lee el Excel, sirve el frontend y habla con Ollama
└── static/
    ├── index.html        # Frontend del chat
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

Puedes editar filas directamente en Excel y reiniciar `server.py`. Si agregas una opción nueva en `Preguntas`,
debe existir en la columna correspondiente de `Outfits`.

## API
- `POST /api/siguiente` `{respuestas}` → próxima pregunta y opciones válidas
- `POST /api/recomendar` `{respuestas}` → outfit
- `POST /api/chat` `{respuestas, outfit, mensaje, historial}` → respuesta de Ollama

## Créditos
Creado por **Laura Galindo, Valentina Rodríguez, Sara Cárdenas y Joseth Sierra** — Grado 1002.
