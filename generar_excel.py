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

# ---------------------------------------------------------------------------
# Estatura, gustos, tiendas y productos (DATOS SINTÉTICOS / FICTICIOS)
# ---------------------------------------------------------------------------
ESTATURAS = [
    # rango, min_cm, max_cm, consejo general, consejo mujer, consejo hombre, evitar
    ("Baja", 0, 154,
     "Alarga la silueta: looks de un solo color, pantalón de tiro alto, rayas verticales y escote en V. Prendas a tu medida, sin exceso de volumen.",
     "Faldas cortas o midi hasta la rodilla, vestidos cruzados, blusas por dentro y calzado con plataforma o del mismo tono del pantalón.",
     "Pantalón slim o recto sin arrugas en el tobillo, camisa por dentro, chaquetas cortas a la cadera y tenis de suela gruesa.",
     "Oversize exagerado, pantalón capri, abrigos muy largos y cortes horizontales marcados."),
    ("Media baja", 155, 164,
     "Juega con el tiro alto y las proporciones: parte de abajo más larga que la de arriba. Los colores similares arriba y abajo te estilizan.",
     "Jeans mom o rectos de tiro alto, faldas midi con botín o tacón medio, crop tops con pantalón alto.",
     "Chinos ajustados, camisetas a la cadera, bomber cortas y mocasines o tenis blancos.",
     "Prendas que te corten a media pierna y chaquetas demasiado largas."),
    ("Media", 165, 174,
     "Tienes una estatura versátil: puedes usar casi cualquier corte. Juega con capas, largos y contrastes de color.",
     "Vestidos midi o maxi, pantalón palazzo o wide leg, blazers largos.",
     "Pantalón recto o cargo, sobrecamisas, gabardinas a media pierna.",
     "Solo cuida que el largo del pantalón no arrugue en exceso sobre el zapato."),
    ("Alta", 175, 184,
     "Aprovecha los volúmenes: oversize, capas, estampados grandes, pantalones anchos y abrigos largos. Un cinturón ayuda a dividir la silueta.",
     "Maxi vestidos, palazzo, faldas largas, calzado plano o de tacón bajo.",
     "Pantalón ancho o cargo, abrigos largos, camisas oversize y botas.",
     "Prendas demasiado cortas o mangas que no llegan a la muñeca: busca tallas 'largo' (tall)."),
    ("Muy alta", 185, 250,
     "Busca tallas 'tall' para mangas y pantalón. Luces muy bien con prendas largas, capas y contrastes horizontales (camisa y pantalón de distinto color).",
     "Vestidos maxi, pantalón wide leg de largo completo, blazers oversize.",
     "Abrigos largos, pantalones de corte recto con largo extra, suéteres gruesos.",
     "Prendas cortas de talla estándar y pantalones que dejen ver el tobillo sin intención."),
]

ESTILOS = [
    # estilo, palabras clave (para reconocerlo en el chat), prendas clave, colores, descripcion
    ("Urbano", "urbano|streetwear|street|calle|hip hop|skater|oversize", "hoodie oversize, jean cargo o jogger, tenis de caña alta, gorra", "negro, gris, blanco, verde militar", "Cómodo, relajado y con actitud."),
    ("Clásico", "clasico|formal|tradicional|elegante sobrio|sobrio", "camisa blanca, pantalón de tela, blazer azul marino, mocasines", "azul marino, beige, blanco, gris", "Prendas atemporales que nunca pasan de moda."),
    ("Deportivo", "deportivo|sport|gym|atletico|athleisure|comodo", "sudadera, leggings o jogger, camiseta dry-fit, tenis running", "negro, blanco, colores neón", "Funcional y cómodo para moverse."),
    ("Bohemio", "bohemio|boho|hippie|hippy|etnico|artesanal", "blusa suelta, falda larga estampada, sandalias, accesorios artesanales", "terracota, mostaza, beige, verde oliva", "Relajado, con texturas y estampados naturales."),
    ("Romántico", "romantico|delicado|femenino|tierno|coqueto", "blusa con volantes, falda plisada, vestido floral, balerinas", "rosa, lila, blanco, pastel", "Suave, con flores, encajes y volantes."),
    ("Minimalista", "minimalista|simple|basico|sencillo|neutro", "camiseta básica, pantalón recto, tenis blancos, bolso liso", "blanco, negro, beige, gris", "Pocas prendas, cortes limpios y colores neutros."),
    ("Elegante", "elegante|fiesta|glamour|sofisticado|de gala", "vestido o traje entallado, tacones o zapatos de cuero, accesorio brillante", "negro, dorado, vinotinto, azul noche", "Sofisticado, ideal para noches y eventos."),
    ("Rockero", "rock|rockero|punk|metal|grunge|dark", "chaqueta de cuero, camiseta de banda, jean roto, botas", "negro, gris oscuro, rojo", "Rebelde, con cuero, taches y botas."),
]

# Tiendas ficticias en Bogotá y Sabana Centro (nombres, direcciones y teléfonos inventados)
TIENDAS = [
    # id, nombre, tipo, publico, estilos, precio, direccion, municipio, lun-vie, sabado, domingo/festivo, calificacion
    ("T01", "Moda Aurora", "Ropa", "Mujer", "Romántico|Clásico|Elegante", "$$", "C.C. Plaza Andina, local 1-24, Calle 80 # 20-15", "Bogotá", "10:00-20:00", "10:00-21:00", "11:00-19:00", 4.6),
    ("T02", "Urbana Store", "Ropa", "Unisex", "Urbano|Minimalista", "$", "Carrera 15 # 45-30", "Bogotá", "09:00-19:00", "09:00-20:00", "Cerrado", 4.4),
    ("T03", "Paso Firme Calzado", "Calzado", "Unisex", "Clásico|Urbano|Elegante", "$$", "C.C. Plaza Andina, local 2-08, Calle 80 # 20-15", "Bogotá", "10:00-20:00", "10:00-21:00", "11:00-19:00", 4.7),
    ("T04", "Detalles & Brillos", "Accesorios", "Unisex", "Romántico|Elegante|Bohemio", "$", "Calle 12 # 5-40, Centro", "Zipaquirá", "08:30-18:30", "08:30-19:00", "09:00-14:00", 4.5),
    ("T05", "Sastrería Don Emilio", "Formal", "Hombre", "Clásico|Elegante", "$$$", "Avenida 19 # 110-22", "Bogotá", "09:00-18:00", "09:00-14:00", "Cerrado", 4.9),
    ("T06", "Activa Sport", "Deportiva", "Unisex", "Deportivo|Urbano", "$$", "C.C. Sabana Norte, local 115, Autopista Norte km 20", "Chía", "10:00-21:00", "10:00-21:00", "10:00-20:00", 4.3),
    ("T07", "Pequeños Pasos", "Infantil", "Infantil", "Clásico|Deportivo|Romántico", "$", "Carrera 3 # 4-18, Parque principal", "Gachancipá", "08:00-19:00", "08:00-19:00", "09:00-13:00", 4.6),
    ("T08", "Boho Luna", "Ropa", "Mujer", "Bohemio|Romántico", "$$", "Calle 11 # 9-35", "Chía", "09:30-19:30", "09:30-20:00", "10:00-17:00", 4.5),
    ("T09", "Rock & Cuero", "Ropa", "Unisex", "Rockero|Urbano", "$$", "Calle 19 # 4-72, local 3", "Bogotá", "10:00-19:00", "10:00-20:00", "Cerrado", 4.2),
    ("T10", "Básicos Nube", "Ropa", "Unisex", "Minimalista|Clásico", "$", "C.C. Sabana Norte, local 204, Autopista Norte km 20", "Chía", "10:00-21:00", "10:00-21:00", "10:00-20:00", 4.4),
    ("T11", "Gala Boutique", "Formal", "Mujer", "Elegante|Clásico", "$$$", "Carrera 11 # 93-40", "Bogotá", "10:00-19:00", "10:00-18:00", "Cerrado", 4.8),
    ("T12", "Tenis Planeta", "Calzado", "Unisex", "Urbano|Deportivo|Minimalista", "$$", "Calle 8 # 6-12", "Zipaquirá", "09:00-19:00", "09:00-20:00", "10:00-16:00", 4.4),
    ("T13", "Mercado de la Moda", "Ropa", "Unisex", "Urbano|Minimalista|Clásico|Romántico", "$", "Carrera 4 # 2-55", "Tocancipá", "08:00-18:00", "08:00-19:00", "09:00-15:00", 4.1),
    ("T14", "Ópticas & Estilo", "Accesorios", "Unisex", "Clásico|Minimalista|Urbano", "$$", "C.C. Plaza Andina, local 1-50, Calle 80 # 20-15", "Bogotá", "10:00-20:00", "10:00-21:00", "11:00-19:00", 4.6),
    ("T15", "Abrigo Andino", "Ropa", "Unisex", "Clásico|Bohemio|Minimalista", "$$$", "Carrera 7 # 72-10", "Bogotá", "09:00-19:00", "10:00-18:00", "Cerrado", 4.7),
]

CATEGORIAS = {  # categoria: [(producto, publico, estilos, precio_base_cop)]
    "Superior": [("Camiseta básica algodón", "Unisex", "Minimalista|Urbano|Deportivo", 35000),
                 ("Hoodie oversize", "Unisex", "Urbano|Deportivo", 95000),
                 ("Camisa de botones", "Hombre", "Clásico|Minimalista", 89000),
                 ("Blusa con volantes", "Mujer", "Romántico", 79000),
                 ("Blusa bohemia bordada", "Mujer", "Bohemio", 85000),
                 ("Camiseta de banda", "Unisex", "Rockero", 55000),
                 ("Top satinado", "Mujer", "Elegante", 99000),
                 ("Suéter de lana", "Unisex", "Clásico|Minimalista|Bohemio", 120000)],
    "Inferior": [("Jean recto", "Unisex", "Minimalista|Clásico|Urbano", 110000),
                 ("Jean mom tiro alto", "Mujer", "Urbano|Romántico|Minimalista", 115000),
                 ("Pantalón cargo", "Unisex", "Urbano", 105000),
                 ("Jogger deportivo", "Unisex", "Deportivo|Urbano", 75000),
                 ("Pantalón chino", "Hombre", "Clásico|Minimalista", 95000),
                 ("Falda midi plisada", "Mujer", "Romántico|Clásico", 89000),
                 ("Falda larga estampada", "Mujer", "Bohemio", 92000),
                 ("Leggings deportivos", "Mujer", "Deportivo", 65000),
                 ("Jean roto", "Unisex", "Rockero|Urbano", 99000)],
    "Vestido/Traje": [("Vestido floral", "Mujer", "Romántico|Bohemio", 139000),
                      ("Vestido de cóctel", "Mujer", "Elegante", 229000),
                      ("Traje de dos piezas", "Hombre", "Clásico|Elegante", 459000),
                      ("Blazer entallado", "Unisex", "Clásico|Elegante|Minimalista", 219000)],
    "Calzado": [("Tenis básicos de cuero", "Unisex", "Minimalista|Urbano|Clásico", 149000),
                ("Tenis de caña alta", "Unisex", "Urbano|Rockero", 179000),
                ("Tenis running", "Unisex", "Deportivo", 199000),
                ("Mocasines de cuero", "Unisex", "Clásico|Elegante", 189000),
                ("Botines", "Unisex", "Rockero|Bohemio|Clásico", 199000),
                ("Tacones", "Mujer", "Elegante", 169000),
                ("Balerinas", "Mujer", "Romántico|Minimalista", 99000),
                ("Sandalias de cuero", "Unisex", "Bohemio", 89000),
                ("Zapatos Oxford", "Hombre", "Clásico|Elegante", 229000)],
    "Accesorio": [("Gorra", "Unisex", "Urbano|Deportivo", 45000),
                  ("Bolso cruzado", "Mujer", "Minimalista|Urbano|Clásico", 89000),
                  ("Clutch brillante", "Mujer", "Elegante", 79000),
                  ("Reloj clásico", "Unisex", "Clásico|Elegante|Minimalista", 159000),
                  ("Gafas de sol", "Unisex", "Urbano|Clásico|Minimalista", 99000),
                  ("Aretes artesanales", "Mujer", "Bohemio|Romántico", 35000),
                  ("Cinturón de cuero", "Unisex", "Clásico|Rockero|Minimalista", 65000),
                  ("Pulsera de taches", "Unisex", "Rockero", 29000),
                  ("Morral", "Unisex", "Urbano|Deportivo|Minimalista", 119000)],
    "Abrigo": [("Chaqueta de jean", "Unisex", "Urbano|Rockero|Minimalista", 139000),
               ("Chaqueta de cuero", "Unisex", "Rockero|Elegante", 349000),
               ("Gabardina", "Unisex", "Clásico|Elegante|Minimalista", 289000),
               ("Chaqueta impermeable", "Unisex", "Deportivo|Urbano", 179000),
               ("Ruana de lana", "Unisex", "Bohemio|Clásico", 149000)],
}
TIPO_CATEGORIAS = {
    "Ropa": ["Superior", "Inferior", "Vestido/Traje", "Abrigo"],
    "Calzado": ["Calzado"],
    "Accesorios": ["Accesorio"],
    "Formal": ["Vestido/Traje", "Superior", "Calzado", "Accesorio", "Abrigo"],
    "Deportiva": ["Superior", "Inferior", "Calzado", "Accesorio", "Abrigo"],
    "Infantil": ["Superior", "Inferior", "Vestido/Traje", "Calzado", "Accesorio", "Abrigo"],
}
FACTOR_PRECIO = {"$": 0.7, "$$": 1.0, "$$$": 1.6}
COLORES = ["negro", "blanco", "azul", "beige", "gris", "rojo", "rosa", "verde", "café", "vinotinto", "lila", "mostaza"]
TALLAS = {"Calzado": "35-43", "Infantil": "2-14", "Accesorio": "Única"}


def filas_productos():
    import random
    rnd = random.Random(1002)  # semilla fija: siempre los mismos datos
    n = 0
    for t in TIENDAS:
        tid, _, tipo, publico, estilos_t, precio = t[0], t[1], t[2], t[3], set(t[4].split("|")), t[5]
        for cat in TIPO_CATEGORIAS[tipo]:
            for prod, pub, estilos_p, base in CATEGORIAS[cat]:
                comunes = estilos_t & set(estilos_p.split("|"))
                if not comunes:
                    continue
                if publico not in ("Unisex", "Infantil") and pub not in ("Unisex", publico):
                    continue
                if tipo == "Infantil" and cat == "Vestido/Traje" and "Traje" in prod:
                    continue
                valor = base * FACTOR_PRECIO[precio] * (0.6 if tipo == "Infantil" else 1) * rnd.uniform(0.9, 1.15)
                valor = int(round(valor / 1000) * 1000 - 100)  # precio tipo $89.900
                nombre = prod + (" infantil" if tipo == "Infantil" else "")
                genero = "Infantil" if tipo == "Infantil" else (pub if publico == "Unisex" else publico)
                for color in rnd.sample(COLORES, 2):
                    n += 1
                    tallas = TALLAS.get("Infantil" if tipo == "Infantil" else cat, "XS-XL")
                    yield [f"PR{n:04d}", tid, nombre, cat, genero, "|".join(sorted(comunes)), color,
                           valor, tallas, rnd.choice(["Sí", "Sí", "Sí", "Pocas unidades"])]


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
hoja(wb, "Estaturas", ["rango", "min_cm", "max_cm", "consejo", "consejo_mujer", "consejo_hombre", "evitar"], ESTATURAS)
hoja(wb, "Estilos", ["estilo", "palabras_clave", "prendas_clave", "colores", "descripcion"], ESTILOS)
hoja(wb, "Tiendas", ["id", "nombre", "tipo", "publico", "estilos", "rango_precio", "direccion", "municipio",
                     "horario_lun_vie", "horario_sabado", "horario_domingo_festivo", "calificacion"],
     TIENDAS)
hoja(wb, "Productos", ["id", "tienda_id", "producto", "categoria", "genero", "estilos", "color", "precio_cop",
                       "tallas", "disponible"], filas_productos())
hoja(wb, "LEEME", ["aviso"], [
    ("⚠️ DATOS SINTÉTICOS: las tiendas, direcciones, teléfonos, horarios y precios de las hojas Tiendas y Productos son ficticios.",),
    ("Fueron creados solo para el proyecto escolar Ropita (Grado 1002). No corresponden a negocios reales.",),
])
wb["LEEME"]["A2"].font = Font(bold=True, color="B91C1C")
wb.move_sheet("LEEME", offset=-(len(wb.sheetnames) - 1))

wb.save("ropita_db.xlsx")
print("Outfits:", wb["Outfits"].max_row - 1, "filas")
