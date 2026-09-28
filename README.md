# Ropita · asistente de estilo

Chatbot que recomienda outfits según género, edad, ocasión, clima, día/noche, día de la semana, **estatura y gustos**,
y sugiere **tiendas con precios, dirección, horario y si están abiertas ahora** (datos ficticios).

👉 **Pruébalo en línea:** https://javier-galindobernal.github.io/ropita/
(en línea funciona todo menos el chat libre con IA, que requiere ejecutarlo en tu computador con Ollama)

## Estructura
```
ropita/
├── ropita_db.xlsx        # Base de conocimiento (editable en Excel)
├── generar_excel.py     # Regenera el Excel desde cero
├── server.py            # Backend: exporta el Excel a JSON, sirve docs/ y habla con Ollama
└── docs/                 # Frontend (publicado en GitHub Pages)
    ├── index.html        # Pantalla del chat
    ├── app.js            # Lógica del chat (reglas, tiendas, horarios)
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

## Qué le puedes escribir
- "Mido 1.60, ¿qué outfit me recomiendas?" → consejos para tu estatura + outfit sugerido
- "Soy mujer y me gusta el estilo bohemio / urbano / clásico…" → prendas y colores de ese estilo
- "Quiero comprar tenis negros por menos de 200 mil" → productos, precios y tiendas
- "¿Qué tiendas están abiertas en Chía?" → horario en tiempo real (hora de Colombia)
- "¿Dónde compro este outfit?" (después del flujo guiado)

## Cómo funciona (modo híbrido)
1. **Preguntas predeterminadas** (hoja `Preguntas`): el bot pregunta con botones. Solo muestra opciones que
   existen en la hoja `Outfits` (p. ej. a un niño no le ofrece "Trabajo").
2. **Recomendación exacta** desde la hoja `Outfits` (sin IA, rápida y controlada).
3. **Chat con reglas** ([docs/app.js](docs/app.js)): reconoce estatura, género, estilos, colores, prendas,
   presupuesto, municipio y horarios, y responde con los datos del Excel. Funciona sin internet ni servidor.
4. **Chat libre con Ollama** (solo en local): lo que las reglas no entienden (p. ej. "¿cómo lavo el cuero?") va a
   Ollama junto con el perfil, el outfit y las tiendas que encajan, para que la IA no invente fuera de la base.

## Hojas del Excel
| Hoja | Contenido |
|---|---|
| Outfits | 1.856 combinaciones: género × edad × ocasión × clima × día/noche × día de la semana |
| Preguntas | Orden, clave, texto y opciones (separadas por `\|`) |
| Prendas | Catálogo de prendas con formalidad, climas y ocasiones |
| Estaturas | 5 rangos de estatura con consejos para mujer, hombre y qué evitar |
| Estilos | 8 estilos (urbano, clásico, bohemio…) con palabras clave, prendas y colores |
| Tiendas | 15 tiendas **ficticias** en Bogotá y Sabana Centro: dirección, horarios, precio, calificación |
| Productos | 372 productos **ficticios** con precio en pesos, color, estilo y tallas |
| Climas / Ocasiones / Grupos_edad | Tablas de referencia |
| LEEME | Aviso: tiendas, direcciones, horarios y precios son datos sintéticos |

Puedes editar filas directamente en Excel y reiniciar `server.py` (regenera `docs/ropita_data.json`).
Para que el cambio se vea en la versión en línea, haz commit y push de `ropita_db.xlsx` y `docs/ropita_data.json`. Si agregas una opción nueva en `Preguntas`,
debe existir en la columna correspondiente de `Outfits`.

## API
- `POST /api/chat` `{respuestas, outfit, mensaje, historial}` → respuesta de Ollama (solo en local)

## Créditos
Creado por **Laura Galindo, Valentina Rodríguez, Sara Cárdenas y Joseth Sierra** — Grado 1002.
