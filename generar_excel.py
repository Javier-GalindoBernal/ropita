"""Genera la base de conocimiento de Ropita (ropita_db.xlsx)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

GENEROS = ["Hombre", "Mujer"]
EDADES = {  # grupo: (rango, etiqueta por género)
    "Infantil": ("3-12 años", {"Hombre": "Niño", "Mujer": "Niña"}),
    "Adolescente": ("13-17 años", {"Hombre": "Adolescente", "Mujer": "Adolescente"}),
    "Joven adulto": ("18-30 años", {"Hombre": "Joven", "Mujer": "Joven"}),
    "Adulto": ("31-55 años", {"Hombre": "Adulto", "Mujer": "Adulta"}),
    "Adulto mayor": ("56+ años", {"Hombre": "Adulto mayor", "Mujer": "Adulta mayor"}),
}
# ocasión: (formalidad, grupos de edad permitidos)
TODOS = list(EDADES)
OCASIONES = {
    "Fiesta": ("Elegante", TODOS),
    "Reunión familiar": ("Smart casual", TODOS),
    "Reunión con amigos": ("Casual", TODOS),
    "Paseo": ("Casual", TODOS),
    "Cita romántica": ("Smart casual", ["Adolescente", "Joven adulto", "Adulto", "Adulto mayor"]),
    "Trabajo": ("Formal", ["Joven adulto", "Adulto", "Adulto mayor"]),
    "Universidad": ("Casual", ["Joven adulto", "Adulto"]),
    "Colegio": ("Casual", ["Infantil", "Adolescente"]),
    "Boda / evento formal": ("Formal", TODOS),
    "Deporte / gimnasio": ("Deportivo", TODOS),
}
CLIMAS = {  # clima: (temperatura, capa, calzado, tela, extra)
    "Caluroso": (">24 °C", "sin capa extra", "sandalias o tenis de lona", "lino y algodón ligero", "gafas de sol y protector solar"),
    "Templado": ("16-24 °C", "chaqueta ligera o cárdigan", "tenis o mocasines", "algodón y denim", "una capa fácil de quitar"),
    "Frío": ("<16 °C", "abrigo o chaqueta acolchada + suéter", "botas o botines cerrados", "lana, franela y punto", "bufanda y medias gruesas"),
    "Lluvioso": ("variable", "chaqueta impermeable o gabardina", "botas impermeables", "telas de secado rápido", "paraguas compacto"),
}
MOMENTOS = {
    "Día": "colores claros y tonos pastel (beige, blanco, celeste)",
    "Noche": "tonos oscuros o intensos (negro, azul marino, vinotinto) con un toque brillante",
}
DIAS = {
    "Entre semana": "prioriza la comodidad y lo práctico",
    "Sábado": "puedes arriesgar más con estampados y accesorios",
    "Domingo": "look relajado y cómodo, ideal para planes tranquilos",
}

# Prendas base por (género, grupo de edad | "*", formalidad): (superior, inferior, calzado, accesorio)
BASE = {
    ("Hombre", "*", "Casual"): ("camiseta o polo", "jean o chino", "tenis blancos", "reloj casual"),
    ("Hombre", "*", "Smart casual"): ("camisa de botones sin corbata", "pantalón chino", "mocasines o zapatos casuales", "correa de cuero"),
    ("Hombre", "*", "Formal"): ("camisa de vestir + blazer", "pantalón de vestir", "zapatos Oxford", "corbata o reloj clásico"),
    ("Hombre", "*", "Elegante"): ("camisa entallada + blazer", "pantalón slim oscuro", "zapatos de cuero o botines", "reloj y perfume"),
    ("Hombre", "*", "Deportivo"): ("camiseta dry-fit", "pantaloneta o jogger", "tenis deportivos", "botella de agua"),
    ("Mujer", "*", "Casual"): ("blusa o camiseta básica", "jean o falda midi", "tenis o sandalias bajas", "bolso cruzado"),
    ("Mujer", "*", "Smart casual"): ("blusa fluida", "pantalón palazzo o falda midi", "balerinas o botines", "aretes discretos"),
    ("Mujer", "*", "Formal"): ("blusa + blazer", "pantalón sastre o falda lápiz", "tacón bajo o stilettos", "bolso estructurado"),
    ("Mujer", "*", "Elegante"): ("vestido de cóctel o top satinado", "(incluido en el vestido) o pantalón de tiro alto", "tacones o sandalias de tacón", "clutch y aretes llamativos"),
    ("Mujer", "*", "Deportivo"): ("top deportivo + camiseta", "leggings o short deportivo", "tenis deportivos", "banda para el cabello"),
    # Niños y niñas
    ("Hombre", "Infantil", "Casual"): ("camiseta con estampado", "jean elástico o bermuda", "tenis con velcro", "gorra"),
    ("Hombre", "Infantil", "Smart casual"): ("camisa tipo polo", "pantalón chino", "tenis limpios o mocasines", "correa infantil"),
    ("Hombre", "Infantil", "Formal"): ("camisa + chaleco", "pantalón de tela", "zapatos de amarrar", "corbatín"),
    ("Hombre", "Infantil", "Elegante"): ("camisa + corbatín", "pantalón de tela", "zapatos de charol", "tirantes"),
    ("Hombre", "Infantil", "Deportivo"): ("camiseta de algodón", "sudadera o pantaloneta", "tenis deportivos", "gorra"),
    ("Mujer", "Infantil", "Casual"): ("camiseta con estampado", "leggings o short", "tenis o sandalias", "diadema o moño"),
    ("Mujer", "Infantil", "Smart casual"): ("blusa con volantes", "falda o jean", "balerinas", "diadema"),
    ("Mujer", "Infantil", "Formal"): ("vestido sencillo", "(incluido en el vestido)", "balerinas", "moño a juego"),
    ("Mujer", "Infantil", "Elegante"): ("vestido de fiesta con tul", "(incluido en el vestido)", "balerinas brillantes", "corona o moño"),
    ("Mujer", "Infantil", "Deportivo"): ("camiseta de algodón", "leggings o sudadera", "tenis deportivos", "cola de caballo"),
    # Adolescentes
    ("Hombre", "Adolescente", "Casual"): ("hoodie o camiseta oversize", "jean o jogger", "tenis urbanos", "gorra o cadena"),
    ("Hombre", "Adolescente", "Elegante"): ("camisa + chaqueta bomber", "jean negro", "tenis blancos limpios", "reloj"),
    ("Mujer", "Adolescente", "Casual"): ("crop top o camiseta oversize", "jean mom o cargo", "tenis urbanos", "mochila pequeña"),
    ("Mujer", "Adolescente", "Elegante"): ("vestido corto o top brillante", "falda o jean de tiro alto", "botines o plataformas", "accesorios dorados"),
    # Adultos mayores
    ("Hombre", "Adulto mayor", "Casual"): ("camisa de manga larga o polo", "pantalón de lino o dril", "zapatos cómodos antideslizantes", "sombrero"),
    ("Mujer", "Adulto mayor", "Casual"): ("blusa holgada", "pantalón de tela fluida", "zapatos planos antideslizantes", "pañoleta"),
    ("Mujer", "Adulto mayor", "Elegante"): ("vestido o conjunto de dos piezas", "pantalón fluido", "tacón bajo cómodo", "collar de perlas"),
}


def base_para(g, e, f):
    return BASE.get((g, e, f)) or BASE[(g, "*", f)]


def tip(clima, momento, dia, ocasion, e):
    t = [f"Clima {clima.lower()}: usa {CLIMAS[clima][3]}; lleva {CLIMAS[clima][4]}."]
    t.append(f"De {momento.lower()}: {MOMENTOS[momento]}.")
    t.append(f"{dia}: {DIAS[dia]}.")
    if e == "Infantil":
        t.append("Prioriza prendas fáciles de poner y una muda extra.")
    if e == "Adulto mayor":
        t.append("Elige calzado con buen agarre y prendas sin ajustes incómodos.")
    if ocasion == "Boda / evento formal" and momento == "Día":
        t.append("Evita el blanco: se reserva para la novia.")
    return " ".join(t)


def estilo(f, momento):
    if f == "Deportivo":
        return "tonos neón o negro técnico" if momento == "Noche" else "colores vivos y transpirables"
    return MOMENTOS[momento].split(" (")[0]


def filas_outfits():
    i = 0
    for g in GENEROS:
        for e, (rango, etiquetas) in EDADES.items():
            for oc, (f, permitidos) in OCASIONES.items():
                if e not in permitidos:
                    continue
                for cl, (temp, capa, calzado_clima, _, _) in CLIMAS.items():
                    for mo in MOMENTOS:
                        for dia in DIAS:
                            if oc in ("Trabajo", "Universidad", "Colegio") and dia == "Domingo":
                                continue
                            sup, inf, calz, acc = base_para(g, e, f)
                            if cl in ("Frío", "Lluvioso") and f in ("Casual", "Deportivo"):
                                calz = calzado_clima
                            if cl in ("Frío", "Lluvioso") and f in ("Elegante", "Formal"):
                                capa = "abrigo largo o gabardina elegante"
                                if g == "Mujer" and e != "Infantil":
                                    calz = "botines de tacón o zapatos cerrados"
                            if cl == "Caluroso" and f == "Casual":
                                inf = inf.replace("jean", "bermuda o short de lino") if g == "Hombre" else inf
                            i += 1
                            yield [f"OUT-{i:04d}", g, e, rango, etiquetas[g], oc, f, cl, temp, mo, dia,
                                   sup, inf, calz, capa, acc, estilo(f, mo), tip(cl, mo, dia, oc, e)]


PREGUNTAS = [
    # orden, clave, pregunta, opciones (separadas por |)
    (1, "genero", "¡Hola! Soy Ropita 👗👔 Te ayudo a elegir qué ponerte. ¿Para quién es el outfit?", "Hombre|Mujer"),
    (2, "grupo_edad", "¿En qué rango de edad está?", "|".join(f"{k} ({v[0]})" for k, v in EDADES.items())),
    (3, "ocasion", "¿Para qué ocasión es?", "|".join(OCASIONES)),
    (4, "clima", "¿Cómo está el clima?", "|".join(f"{k} ({v[0]})" for k, v in CLIMAS.items())),
    (5, "momento", "¿Es de día o de noche?", "|".join(MOMENTOS)),
    (6, "tipo_dia", "¿Qué día es?", "|".join(DIAS)),
]

PRENDAS = [
    # id, prenda, categoría, género, formalidad, climas, ocasiones sugeridas
    ("P001", "Camiseta básica", "Superior", "Unisex", "Casual", "Caluroso|Templado", "Paseo|Reunión con amigos|Universidad"),
    ("P002", "Camisa de botones", "Superior", "Hombre", "Smart casual", "Templado|Caluroso", "Reunión familiar|Cita romántica|Trabajo"),
    ("P003", "Blusa fluida", "Superior", "Mujer", "Smart casual", "Caluroso|Templado", "Reunión familiar|Cita romántica"),
    ("P004", "Blazer", "Capa", "Unisex", "Formal", "Templado|Frío", "Trabajo|Boda / evento formal|Fiesta"),
    ("P005", "Hoodie", "Superior", "Unisex", "Casual", "Templado|Frío", "Universidad|Paseo|Colegio"),
    ("P006", "Suéter de lana", "Superior", "Unisex", "Smart casual", "Frío", "Reunión familiar|Trabajo"),
    ("P007", "Ropitado de cóctel", "Pieza completa", "Mujer", "Elegante", "Templado|Caluroso", "Fiesta|Boda / evento formal"),
    ("P008", "Jean", "Inferior", "Unisex", "Casual", "Templado|Frío", "Paseo|Universidad|Reunión con amigos"),
    ("P009", "Pantalón chino", "Inferior", "Unisex", "Smart casual", "Templado|Caluroso", "Reunión familiar|Trabajo|Cita romántica"),
    ("P010", "Pantalón de vestir", "Inferior", "Unisex", "Formal", "Todos", "Trabajo|Boda / evento formal"),
    ("P011", "Falda midi", "Inferior", "Mujer", "Smart casual", "Caluroso|Templado", "Reunión familiar|Cita romántica"),
    ("P012", "Bermuda de lino", "Inferior", "Hombre", "Casual", "Caluroso", "Paseo|Reunión con amigos"),
    ("P013", "Leggings", "Inferior", "Mujer", "Deportivo", "Todos", "Deporte / gimnasio|Paseo"),
    ("P014", "Tenis blancos", "Calzado", "Unisex", "Casual", "Caluroso|Templado", "Paseo|Universidad|Reunión con amigos"),
    ("P015", "Mocasines", "Calzado", "Unisex", "Smart casual", "Templado", "Reunión familiar|Trabajo"),
    ("P016", "Zapatos Oxford", "Calzado", "Hombre", "Formal", "Todos", "Trabajo|Boda / evento formal"),
    ("P017", "Tacones", "Calzado", "Mujer", "Elegante", "Caluroso|Templado", "Fiesta|Boda / evento formal"),
    ("P018", "Botas impermeables", "Calzado", "Unisex", "Casual", "Lluvioso|Frío", "Paseo|Universidad"),
    ("P019", "Abrigo", "Capa", "Unisex", "Smart casual", "Frío", "Trabajo|Reunión familiar|Fiesta"),
    ("P020", "Chaqueta impermeable", "Capa", "Unisex", "Casual", "Lluvioso", "Paseo|Universidad|Colegio"),
    ("P021", "Sandalias", "Calzado", "Unisex", "Casual", "Caluroso", "Paseo|Reunión con amigos"),
    ("P022", "Gafas de sol", "Accesorio", "Unisex", "Casual", "Caluroso", "Paseo"),
    ("P023", "Bufanda", "Accesorio", "Unisex", "Smart casual", "Frío", "Todas"),
    ("P024", "Paraguas", "Accesorio", "Unisex", "Casual", "Lluvioso", "Todas"),
]


def hoja(wb, titulo, encabezados, filas, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws.append(encabezados)
    for f in filas:
        ws.append(list(f))
    fill = PatternFill("solid", fgColor="7C3AED")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for idx, _ in enumerate(encabezados, 1):
        largo = max(len(str(ws.cell(r, idx).value or "")) for r in range(1, min(ws.max_row, 60) + 1))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, largo + 2), 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


wb = Workbook()
hoja(wb, "Outfits",
     ["id", "genero", "grupo_edad", "rango_edad", "perfil", "ocasion", "formalidad", "clima", "temperatura",
      "momento", "tipo_dia", "prenda_superior", "prenda_inferior", "calzado", "capa_abrigo", "accesorio",
      "paleta_colores", "consejo"],
     filas_outfits(), primera=True)
hoja(wb, "Preguntas", ["orden", "clave", "pregunta", "opciones"], PREGUNTAS)
hoja(wb, "Prendas", ["id", "prenda", "categoria", "genero", "formalidad", "climas", "ocasiones"], PRENDAS)
hoja(wb, "Climas", ["clima", "temperatura", "capa", "calzado", "telas", "extra"], [(k, *v) for k, v in CLIMAS.items()])
hoja(wb, "Ocasiones", ["ocasion", "formalidad", "grupos_edad_permitidos"],
     [(k, f, ", ".join(p)) for k, (f, p) in OCASIONES.items()])
hoja(wb, "Grupos_edad", ["grupo_edad", "rango", "etiqueta_hombre", "etiqueta_mujer"],
     [(k, r, et["Hombre"], et["Mujer"]) for k, (r, et) in EDADES.items()])

wb.save("ropita_db.xlsx")
print("Outfits:", wb["Outfits"].max_row - 1, "filas")
