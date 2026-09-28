// Ropita · lógica del chat.
// Todo funciona en el navegador con ropita_data.json (también en GitHub Pages).
// En local (python3 server.py) las preguntas que las reglas no entienden van a Ollama.

const LOCAL = ['localhost', '127.0.0.1'].includes(location.hostname);
const chat = document.getElementById('chat');
const form = document.getElementById('form');
const input = document.getElementById('texto');
const btn = document.getElementById('enviar');

let DB = null;
let respuestas = {}, outfit = null, historial = [];
const perfil = { estatura: null, genero: null, estilos: [], colores: [], presupuesto: null };

// ---------- utilidades ----------
const norm = t => String(t).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const limpiar = v => String(v).replace(/\s*\(.*\)$/, '').trim();  // 'Infantil (3-12 años)' -> 'Infantil'
const pesos = n => new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(n);
const lista = v => String(v || '').split('|').filter(Boolean);

function el(tag, clase, texto) {
  const e = document.createElement(tag);
  if (clase) e.className = clase;
  if (texto != null) e.textContent = texto;
  return e;
}
const bajar = () => { chat.scrollTop = chat.scrollHeight; };

function burbuja(texto, quien = 'bot') {
  const div = el('div', 'msg ' + quien);
  if (quien === 'bot') div.innerHTML = '<img src="ropita-avatar.svg" alt="">';
  const b = el('div', 'burbuja', texto);
  div.appendChild(b); chat.appendChild(div); bajar();
  return b;
}

function botones(opciones, alElegir, clase = '') {
  const cont = el('div', 'opciones ' + clase);
  opciones.forEach(op => {
    const b = el('button', '', op);
    b.type = 'button';
    b.onclick = () => { cont.remove(); alElegir(op); };
    cont.appendChild(b);
  });
  chat.appendChild(cont); bajar();
}

// Tarjeta con título y filas [etiqueta, valor]
function tarjeta(titulo, filas = [], secciones = []) {
  const card = el('div', 'card');
  card.appendChild(el('h3', '', titulo));
  if (filas.length) {
    const dl = el('dl');
    filas.forEach(([k, v]) => dl.append(el('dt', '', k), el('dd', '', v)));
    card.appendChild(dl);
  }
  secciones.forEach(([k, v]) => {
    const s = el('div', 'seccion');
    if (k) s.append(el('b', '', k + ' '));
    s.append(v);
    card.appendChild(s);
  });
  chat.appendChild(card); bajar();
  return card;
}

// ---------- horario de tiendas (hora de Colombia) ----------
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function ahoraBogota() {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Bogota', weekday: 'short', hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
  }).formatToParts(new Date()).map(x => [x.type, x.value]));
  return { dia: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].indexOf(p.weekday), min: +p.hour * 60 + +p.minute };
}

function horarioDelDia(t, dia) {
  const h = dia === 0 ? t.horario_domingo_festivo : dia === 6 ? t.horario_sabado : t.horario_lun_vie;
  if (!h || /cerrad/i.test(h)) return null;
  const [a, c] = h.split('-').map(x => { const [hh, mm] = x.split(':'); return +hh * 60 + +mm; });
  return { abre: a, cierra: c, texto: h };
}
const hhmm = m => `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;

function estadoTienda(t) {
  const { dia, min } = ahoraBogota();
  const hoy = horarioDelDia(t, dia);
  if (hoy && min >= hoy.abre && min < hoy.cierra) return { abierta: true, texto: `Abierta · cierra a las ${hhmm(hoy.cierra)}` };
  if (hoy && min < hoy.abre) return { abierta: false, texto: `Cerrada · abre hoy a las ${hhmm(hoy.abre)}` };
  for (let i = 1; i <= 7; i++) {
    const d = (dia + i) % 7, h = horarioDelDia(t, d);
    if (h) return { abierta: false, texto: `Cerrada · abre ${i === 1 ? 'mañana' : 'el ' + DIAS[d]} a las ${hhmm(h.abre)}` };
  }
  return { abierta: false, texto: 'Cerrada' };
}

// ---------- entender el mensaje ----------
const CATEGORIAS = [
  ['Calzado', /\b(zapat\w*|tenis|botas?|botines?|tacon\w*|sandalias?|calzado|balerinas?|mocasin\w*|oxford)\b/],
  ['Accesorio', /\b(accesorio\w*|bolsos?|carteras?|reloj\w*|gafas|lentes|aretes?|gorras?|cinturon\w*|correas?|pulseras?|morral\w*|collar\w*|clutch)\b/],
  ['Abrigo', /\b(chaquetas?|abrigos?|gabardinas?|ruanas?|chamarras?|chaqueton\w*)\b/],
  ['Superior', /\b(camisas?|camisetas?|blusas?|hoodies?|buzos?|sueter\w*|tops?|polos?)\b/],
  ['Inferior', /\b(pantalon\w*|jeans?|faldas?|joggers?|leggings?|shorts?|sudaderas?)\b/],
  ['Vestido/Traje', /\b(vestidos?|trajes?|blazers?|vestido de fiesta)\b/],
];
const COLORES = [
  ['negro', /\bnegr[oa]s?\b/], ['blanco', /\bblanc[oa]s?\b/], ['azul', /\bazul(es)?\b/], ['beige', /\bbeige\b/],
  ['gris', /\bgris(es)?\b/], ['rojo', /\broj[oa]s?\b/], ['rosa', /\b(rosa|rosad[oa])s?\b/], ['verde', /\bverdes?\b/],
  ['café', /\b(cafe|marron)(es)?\b/], ['vinotinto', /\bvinotinto\b/], ['lila', /\b(lila|morad[oa])s?\b/], ['mostaza', /\bmostaza\b/],
];
const MUNICIPIOS = ['Bogotá', 'Chía', 'Zipaquirá', 'Gachancipá', 'Tocancipá'];

function entender(texto) {
  const t = norm(texto);
  const r = { categorias: [], estilos: [], colores: [] };

  // Estatura: "1.60", "1,60 m", "160 cm", "mido 160"
  let m = t.match(/\b([12])[.,](\d{1,2})\s*(m|mt|mts|metros?|cm)?\b/);
  if (m) r.estatura = +m[1] * 100 + +m[2].padEnd(2, '0');
  else if ((m = t.match(/\b(\d{2,3})\s*(cm|centimetros)\b/)) || (m = t.match(/\bmido\s+(\d{2,3})\b/))) r.estatura = +m[1];
  if (r.estatura && (r.estatura < 50 || r.estatura > 230)) delete r.estatura;

  if (/\b(hombre|chico|masculino|varon|caballero|novio|papa|esposo)\b/.test(t)) r.genero = 'Hombre';
  if (/\b(mujer|chica|femenin[oa]|dama|novia|mama|esposa)\b/.test(t)) r.genero = 'Mujer';
  if (/\b(nin[oa]s?|bebe|infantil|hij[oa])\b/.test(t)) r.genero = 'Infantil';

  DB.estilos.forEach(e => {
    if (lista(e.palabras_clave).concat(norm(e.estilo)).some(k => new RegExp(`\\b${norm(k)}\\b`).test(t))) r.estilos.push(e.estilo);
  });
  COLORES.forEach(([c, re]) => re.test(t) && r.colores.push(c));
  CATEGORIAS.forEach(([c, re]) => re.test(t) && r.categorias.push(c));
  if (/\bropa\b/.test(t) && !r.categorias.length) r.categorias.push('Superior', 'Inferior', 'Vestido/Traje', 'Abrigo');
  r.municipio = MUNICIPIOS.find(x => t.includes(norm(x)));

  // Presupuesto: "menos de 100 mil", "hasta $150.000", "presupuesto 200000", "80k"
  m = t.match(/(?:menos de|hasta|maximo|max|presupuesto(?: de)?|tengo|con)\s*\$?\s*(\d[\d.]*)\s*(mil|k|pesos)?/) ||
      t.match(/\$\s*(\d[\d.]*)\s*(mil|k)?/) || t.match(/\b(\d[\d.]*)\s*(mil|k|pesos)\b/);
  if (m) {
    const v = +m[1].replace(/\./g, '') * (/mil|k/.test(m[2] || '') ? 1000 : 1);
    if (v >= 5000) r.presupuesto = v;
  }
  if (/\b(barat\w*|economic\w*|poca plata|no muy caro)\b/.test(t)) r.barato = true;

  r.compra = /\b(compr\w*|tiendas?|donde|precios?|cuesta|cuestan|vale|valen|venden|conseguir|consigo|almacen\w*|barat\w*|economic\w*)\b/.test(t) || !!r.presupuesto;
  r.horario = /\b(abiert\w*|abre|abren|cierra|cierran|horarios?|atencion|a que hora)\b/.test(t);
  r.outfit = /\b(outfit\w*|recomiend\w*|recomendar\w*|que me pongo|ponerme|me visto|vestirme|luzco|favorece\w*|combina\w*|look)\b/.test(t);
  // Preguntas de "cómo / por qué" (lavar, cuidar…) no son compras aunque nombren una prenda
  r.pregunta = /\b(como|por que|porque|cuando|que es|lav\w*|planch\w*|cuid\w*|guard\w*|doblar|manch\w*)\b/.test(t);
  r.esteOutfit = /\beste (outfit|look)\b|\blo compro\b/.test(t);
  r.saludo = /^(hola|buenas|buenos dias|buenas tardes|buenas noches|hey|ayuda|que puedes hacer)\b/.test(t.trim());
  r.gracias = /\b(gracias|genial|excelente|perfecto|chevere)\b/.test(t);
  return r;
}

// ---------- respuestas ----------
function rangoEstatura(cm) {
  return DB.estaturas.find(e => cm >= e.min_cm && cm <= e.max_cm) || DB.estaturas.at(-1);
}

function responderEstatura(cm) {
  const e = rangoEstatura(cm), g = perfil.genero;
  const secciones = [['✨ Tip:', e.consejo]];
  if (g === 'Mujer' || !g || g === 'Infantil') secciones.push(['👩 Para mujer:', e.consejo_mujer]);
  if (g === 'Hombre' || !g || g === 'Infantil') secciones.push(['👨 Para hombre:', e.consejo_hombre]);
  secciones.push(['🚫 Mejor evita:', e.evitar]);
  if (cm < 140) secciones.push(['🧒 Niños:', 'A esta edad se crece rápido: elige tallas con un poco de margen y telas elásticas.']);
  tarjeta(`📏 ${(cm / 100).toFixed(2)} m · estatura ${e.rango.toLowerCase()}`, [], secciones);
}

function estiloActual() {
  return DB.estilos.find(e => e.estilo === perfil.estilos[0]) || DB.estilos.find(e => e.estilo === 'Minimalista');
}

function responderOutfitPersonal() {
  const e = estiloActual();
  const filas = [['🧩 Prendas', e.prendas_clave], ['🎨 Colores', perfil.colores.length ? perfil.colores.join(', ') : e.colores]];
  const secciones = [['💬', e.descripcion]];
  if (perfil.estatura) {
    const r = rangoEstatura(perfil.estatura);
    const tip = perfil.genero === 'Hombre' ? r.consejo_hombre : perfil.genero === 'Mujer' ? r.consejo_mujer : r.consejo;
    secciones.push([`📏 Ajuste por tu estatura (${(perfil.estatura / 100).toFixed(2)} m):`, tip]);
  }
  tarjeta(`👗 Outfit sugerido · estilo ${e.estilo.toLowerCase()}`, filas, secciones);
  if (!perfil.estilos.length) burbuja('Cuéntame tus gustos (urbano, clásico, romántico, bohemio, deportivo, minimalista, elegante o rockero) y lo ajusto. 😉');
}

function responderEstilos(estilos) {
  estilos.forEach(n => {
    const e = DB.estilos.find(x => x.estilo === n);
    tarjeta(`💜 Estilo ${e.estilo.toLowerCase()}`, [['🧩 Prendas clave', e.prendas_clave], ['🎨 Colores', e.colores]], [['💬', e.descripcion]]);
  });
}

const ESTILOS_POR_FORMALIDAD = {
  'Elegante': ['Elegante'], 'Formal': ['Clásico', 'Elegante'], 'Smart casual': ['Clásico', 'Minimalista', 'Romántico'],
  'Casual': ['Urbano', 'Minimalista'], 'Deportivo': ['Deportivo'],
};

// Busca productos y los agrupa por tienda (máx. 3 tiendas × 3 productos)
function buscar({ categorias = [], estilos = [], colores = [], genero = null, presupuesto = null, barato = false, municipio = null }) {
  const tiendas = Object.fromEntries(DB.tiendas.map(t => [t.id, t]));
  let productos = DB.productos.filter(p => {
    const t = tiendas[p.tienda_id];
    if (categorias.length && !categorias.includes(p.categoria)) return false;
    if (presupuesto && p.precio_cop > presupuesto) return false;
    if (municipio && t.municipio !== municipio) return false;
    if (genero === 'Infantil') return p.genero === 'Infantil';
    if (p.genero === 'Infantil') return false;
    return !genero || p.genero === genero || p.genero === 'Unisex';
  });
  // Si hay productos del color / estilo pedido, se muestran solo esos
  const conColor = productos.filter(p => colores.includes(p.color));
  if (conColor.length) productos = conColor;
  const conEstilo = productos.filter(p => lista(p.estilos).some(x => estilos.includes(x)));
  if (conEstilo.length) productos = conEstilo;
  const puntaje = p => {
    let s = 0;
    if (estilos.length && lista(p.estilos).some(x => estilos.includes(x))) s += 3;
    if (colores.includes(p.color)) s += 2;
    if (p.disponible === 'Sí') s += 0.5;
    if (estadoTienda(tiendas[p.tienda_id]).abierta) s += 0.5;
    return s;
  };
  productos = productos.map(p => ({ ...p, _s: puntaje(p) }))
    .sort((a, b) => b._s - a._s || (barato ? a.precio_cop - b.precio_cop : tiendas[b.tienda_id].calificacion - tiendas[a.tienda_id].calificacion));
  const grupos = [];
  for (const p of productos) {
    let g = grupos.find(g => g.tienda.id === p.tienda_id);
    if (!g) { if (grupos.length === 3) continue; g = { tienda: tiendas[p.tienda_id], productos: [] }; grupos.push(g); }
    const repetido = g.productos.some(x => (x.producto === p.producto && x.color === p.color) ||
                                           (categorias.length > 1 && x.categoria === p.categoria));  // variedad de prendas
    if (g.productos.length < 3 && !repetido) g.productos.push(p);
  }
  return { total: productos.length, grupos };
}

function tarjetaTienda(t, productos = []) {
  const est = estadoTienda(t);
  const card = el('div', 'card tienda');
  const h = el('h3');
  h.append(el('span', '', `🏬 ${t.nombre}`), el('span', '', `⭐ ${t.calificacion} · ${t.rango_precio}`));
  card.appendChild(h);
  const meta = el('div', 'meta');
  meta.append(el('span', 'badge ' + (est.abierta ? 'si' : 'no'), (est.abierta ? '🟢 ' : '🔴 ') + est.texto), el('br'),
              `📍 ${t.direccion}, ${t.municipio}`, el('br'),
              `🕘 L-V ${t.horario_lun_vie} · Sáb ${t.horario_sabado} · Dom/fest ${t.horario_domingo_festivo}`);
  card.appendChild(meta);
  if (productos.length) {
    const ul = el('ul');
    productos.forEach(p => {
      const li = el('li'), izq = el('span', '', `${p.producto} · ${p.color}`);
      izq.append(el('br'), el('small', '', `Tallas ${p.tallas}${p.disponible !== 'Sí' ? ' · ' + p.disponible.toLowerCase() : ''}`));
      li.append(izq, el('span', 'precio', pesos(p.precio_cop)));
      ul.appendChild(li);
    });
    card.appendChild(ul);
  }
  chat.appendChild(card); bajar();
}

const notaFicticia = () => { chat.appendChild(el('div', 'nota', 'ℹ️ Tiendas, direcciones, horarios y precios son datos ficticios del proyecto escolar.')); bajar(); };

function responderCompra(r) {
  let filtro = { categorias: r.categorias, estilos: perfil.estilos, colores: perfil.colores, genero: perfil.genero,
                 presupuesto: perfil.presupuesto, barato: r.barato, municipio: r.municipio };
  if (r.esteOutfit && outfit && !r.categorias.length) {
    filtro.categorias = ['Superior', 'Inferior', 'Calzado', 'Vestido/Traje'];
    filtro.estilos = ESTILOS_POR_FORMALIDAD[outfit.formalidad] || [];
    filtro.colores = [];  // para el outfit completo no se fuerza un color
  }
  const { total, grupos } = buscar(filtro);
  if (!grupos.length) {
    burbuja(`No encontré productos con esos filtros${perfil.presupuesto ? ` por menos de ${pesos(perfil.presupuesto)}` : ''}. Prueba con otro presupuesto, color o municipio.`);
    return;
  }
  const que = filtro.categorias.length ? filtro.categorias.map(c => c.toLowerCase()).join(', ') : 'productos';
  burbuja(`🛍️ Encontré ${total} ${total === 1 ? 'opción' : 'opciones'} de ${que}${perfil.presupuesto ? ` por menos de ${pesos(perfil.presupuesto)}` : ''}. Estas son las mejores para ti:`);
  grupos.forEach(g => tarjetaTienda(g.tienda, g.productos));
  notaFicticia();
}

function responderHorario(r) {
  let tiendas = DB.tiendas.filter(t => !r.municipio || t.municipio === r.municipio);
  if (r.categorias.length) {
    const conCategoria = new Set(DB.productos.filter(p => r.categorias.includes(p.categoria)).map(p => p.tienda_id));
    tiendas = tiendas.filter(t => conCategoria.has(t.id));
  }
  const abiertas = tiendas.filter(t => estadoTienda(t).abierta);
  const { dia, min } = ahoraBogota();
  burbuja(`🕘 Hoy es ${DIAS[dia]}, son las ${hhmm(min)} (hora de Colombia). ${abiertas.length ? `Hay ${abiertas.length} tiendas abiertas${r.municipio ? ' en ' + r.municipio : ''}:` : 'Ahora no hay tiendas abiertas; estas abren pronto:'}`);
  (abiertas.length ? abiertas : tiendas).slice(0, 4).forEach(t => tarjetaTienda(t));
  notaFicticia();
}

function ayuda() {
  burbuja('Puedes escribirme cosas como:\n• "Mido 1.60, ¿qué outfit me recomiendas?"\n• "Soy mujer y me gusta el estilo bohemio"\n• "Quiero comprar tenis negros por menos de 150 mil"\n• "¿Qué tiendas están abiertas en Chía?"\nO arma un outfit paso a paso según el clima y la ocasión.');
  chipsInicio();
}

function chipsInicio() {
  botones(['👗 Armar outfit paso a paso', 'Mido 1.60, ¿qué outfit me recomiendas?', 'Me gusta el estilo urbano',
           'Quiero comprar tenis blancos', '¿Qué tiendas están abiertas ahora?'],
          op => op.startsWith('👗') ? (burbuja(op, 'user'), iniciarFlujo()) : procesar(op), 'chips');
}

// Contexto compacto para Ollama con tiendas y productos que encajan con el perfil
function contextoOllama(r) {
  const { grupos } = buscar({ categorias: r.categorias, estilos: perfil.estilos, colores: perfil.colores, genero: perfil.genero, presupuesto: perfil.presupuesto });
  return JSON.stringify(grupos.map(g => ({
    tienda: g.tienda.nombre, direccion: `${g.tienda.direccion}, ${g.tienda.municipio}`, estado: estadoTienda(g.tienda).texto,
    productos: g.productos.map(p => `${p.producto} ${p.color} ${pesos(p.precio_cop)}`),
  })));
}

async function preguntarOllama(mensaje, r) {
  const espera = burbuja('Ropita está pensando…'); espera.classList.add('escribiendo');
  const datos = await fetch('/api/chat', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ respuestas: { ...respuestas, ...perfil }, outfit, mensaje, historial, contexto: contextoOllama(r) }),
  }).then(r => r.json()).catch(() => ({ error: 'No pude conectar con el servidor de Ropita.' }));
  espera.parentElement.remove();
  const texto = datos.respuesta || '⚠️ ' + datos.error;
  burbuja(texto);
  if (datos.respuesta) historial.push({ role: 'user', content: mensaje }, { role: 'assistant', content: texto });
}

async function procesar(mensaje) {
  burbuja(mensaje, 'user');
  const r = entender(mensaje);
  // Recordar lo que la persona nos cuenta
  if (r.estatura) perfil.estatura = r.estatura;
  if (r.genero) perfil.genero = r.genero;
  if (r.estilos.length) perfil.estilos = r.estilos;
  if (r.colores.length) perfil.colores = r.colores;
  if (r.presupuesto) perfil.presupuesto = r.presupuesto;

  let respondio = false;
  if (r.saludo && !r.estatura && !r.compra && !r.estilos.length) { burbuja('¡Hola! 👋'); ayuda(); return; }
  if (r.estatura) { responderEstatura(r.estatura); respondio = true; }
  if (r.estilos.length && !r.compra && !r.outfit && !r.estatura) { responderEstilos(r.estilos); respondio = true; }
  if ((r.outfit || r.estatura) && !r.compra) { responderOutfitPersonal(); respondio = true; }
  if (r.horario && !r.categorias.length && !r.presupuesto) { responderHorario(r); respondio = true; }
  else if (r.compra || (r.categorias.length && !r.pregunta)) { responderCompra(r); respondio = true; }

  if (respondio) {
    if (!r.compra) botones([`🛍️ ¿Dónde compro ropa de estilo ${estiloActual().estilo.toLowerCase()}?`, '👗 Armar outfit por ocasión'],
                           op => op.startsWith('👗') ? (burbuja(op, 'user'), iniciarFlujo()) : procesar(op), 'chips');
    return;
  }
  if (r.gracias) { burbuja('¡Con gusto! 💜 ¿Te ayudo con algo más?'); return; }
  if (LOCAL) return preguntarOllama(mensaje, r);
  burbuja('Mmm, no entendí bien. 🤔 (En esta versión en línea respondo sobre estatura, gustos, outfits y tiendas; el chat libre con IA funciona al ejecutar Ropita en tu computador con Ollama.)');
  ayuda();
}

// ---------- flujo guiado por clima y ocasión ----------
const filtrar = r => DB.outfits.filter(o => Object.entries(r).every(([k, v]) => o[k] === limpiar(v)));

function iniciarFlujo() { respuestas = {}; outfit = null; siguiente(); }

function siguiente() {
  const p = DB.preguntas.find(p => !(p.clave in respuestas));
  if (!p) return recomendar();
  // Solo se ofrecen opciones que tengan al menos un outfit con las respuestas previas
  const posibles = new Set(filtrar(respuestas).map(o => o[p.clave]));
  burbuja(p.pregunta);
  botones(lista(p.opciones).filter(op => posibles.has(limpiar(op))),
          op => { burbuja(op, 'user'); respuestas[p.clave] = op; siguiente(); });
}

function recomendar() {
  outfit = filtrar(respuestas)[0];
  if (!outfit) { burbuja('No encontré un outfit para esa combinación.'); return iniciarFlujo(); }
  perfil.genero = outfit.grupo_edad === 'Infantil' ? 'Infantil' : outfit.genero;
  burbuja(`¡Listo! Este es tu outfit para ${outfit.ocasion.toLowerCase()} ✨`);
  const card = tarjeta(`${outfit.perfil} · ${outfit.formalidad} · ${outfit.clima} · ${outfit.momento}`,
    [['👕 Superior', outfit.prenda_superior], ['👖 Inferior', outfit.prenda_inferior], ['👟 Calzado', outfit.calzado],
     ['🧥 Abrigo', outfit.capa_abrigo], ['💍 Accesorio', outfit.accesorio], ['🎨 Colores', outfit.paleta_colores]]);
  const c = el('div', 'consejo', '💡 ' + outfit.consejo);
  card.appendChild(c);
  if (perfil.estatura) {
    const r = rangoEstatura(perfil.estatura);
    card.appendChild(el('div', 'seccion', `📏 Por tu estatura (${(perfil.estatura / 100).toFixed(2)} m): ${perfil.genero === 'Hombre' ? r.consejo_hombre : r.consejo_mujer}`));
  }
  burbuja('¿Quieres saber dónde comprarlo? También puedes escribirme tu estatura o tus gustos para afinarlo.');
  botones(['🛍️ ¿Dónde compro este outfit?', '🔄 Nuevo outfit'],
          op => op.startsWith('🔄') ? (burbuja(op, 'user'), iniciarFlujo()) : procesar(op), 'chips');
}

// ---------- inicio ----------
form.onsubmit = e => {
  e.preventDefault();
  const mensaje = input.value.trim();
  if (!mensaje || !DB) return;
  input.value = '';
  procesar(mensaje);
};

fetch('ropita_data.json').then(r => r.json()).then(d => {
  DB = d;
  burbuja('¡Hola! Soy Ropita 👗👔 Te ayudo a elegir qué ponerte según el clima, la ocasión, tu estatura y tus gustos, y te digo dónde comprarlo. Escríbeme o elige una opción:');
  chipsInicio();
}).catch(() => burbuja('No pude cargar la base de datos de Ropita.'));
