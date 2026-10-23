const PRINCIPALES = [
  "Folio",
  "Nombre",
  "Alcance",
  "Beneficios esperados",
  "JP",
  "LI",
  "Fecha inicio",
  "Fecha término",
  "Avance Plan",
  "Avance real",
  "Desviación",
  "Días retraso",
  "Actividad semana en curso",
  "Actividades semana próxima",
  "Riesgos",
  "Accionables",
  "Responsable",
  "Etapa",
  "Principales Hitos",
];
let proyectos = [],
  filtrados = [],
  proyectoSeleccionado = null;
const ETAPAS = [
  "Maduración",
  "Construcción",
  "Despliegue",
  "Cierre",
  "Detenido",
];
const ESTADOS = ["Backlog", "Maduración", "En Curso", "Cerrado", "Detenido"];
const HOJA_PRINCIPAL = "Detalle semanal";
const $ = (id) => document.getElementById(id);
const visible = (v) =>
  v === null || v === undefined || String(v).trim() === ""
    ? "-"
    : String(v).trim();

function salud(v) {
  if (v === null || v === undefined || String(v).trim() === "")
    return ["-", "neutral"];
  let s = String(v).replace(",", "."),
    n = Number(s.replace("%", ""));
  if (Number.isNaN(n)) return ["-", "neutral"];
  if (!s.includes("%") && Math.abs(n) <= 1) n *= 100;
  n = Math.abs(n);
  return n < 5
    ? ["Saludable", "green"]
    : n <= 15
      ? ["Desviado", "yellow"]
      : ["Crítico", "red"];
}
function pct(v) {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  let s = String(v).replace(",", ".").trim(),
    n = Number(s.replace("%", ""));
  if (Number.isNaN(n)) return null;
  if (!s.includes("%") && Math.abs(n) <= 1) n *= 100;
  return n;
}
function esc(v) {
  return String(v)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
$("excel").onchange = async (e) => {
  let f = e.target.files[0];
  if (!f) return;
  try {
    $("estado").textContent = "Cargando " + f.name + "...";
    let b = await f.arrayBuffer(),
      w = XLSX.read(b, { type: "array", cellDates: true }),
      ws = w.Sheets[HOJA_PRINCIPAL],
      ws2 = w.Sheets["Proyectos mensual"],
      rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false }),
      rows2 = XLSX.utils.sheet_to_json(ws2, { defval: null, raw: false });
    proyectos = rows.map((r) => {
      const normalizado = {};
      Object.entries(r).forEach(([k, v]) => {
        normalizado[String(k).trim().replace(/\s+/g, " ")] = v;
      });
      normalizado._salud = salud(normalizado["Desviación"]);
      return normalizado;
    });
    filtrados = [...proyectos];
    etapas();
    render();
    $("estado").textContent = f.name + " · " + proyectos.length + " proyectos";
  } catch (err) {
    $("estado").textContent = "Error";
    alert(err.message);
  }
};
function etapas() {
  let vals = [
    ...new Set([
      ...ETAPAS,
      ...proyectos.map((p) => visible(p["Etapa"])).filter((v) => v !== "-"),
    ]),
  ];
  $("etapa").innerHTML =
    '<option value="">Todas las etapas</option>' +
    vals.map((v) => `<option>${esc(v)}</option>`).join("");
}
["buscar", "etapa", "salud"].forEach((id) =>
  $(id).addEventListener("input", filtrar),
);
function filtrar() {
  let q = $("buscar").value.toLowerCase(),
    e = $("etapa").value,
    s = $("salud").value;
  filtrados = proyectos.filter(
    (p) =>
      (!q ||
        [p.Folio, p.Nombre, p.JP, p.LI]
          .map(visible)
          .join(" ")
          .toLowerCase()
          .includes(q)) &&
      (!e || visible(p.Etapa) === e) &&
      (!s || p._salud[0] === s),
  );
  render();
}
function render() {
  let c = { Saludable: 0, Desviado: 0, Crítico: 0 };
  proyectos.forEach((p) => {
    if (c[p._salud[0]] !== undefined) c[p._salud[0]]++;
  });
  $("summary").innerHTML =
    `<div class=sum><b>${proyectos.length}</b><span>Total</span></div><div class=sum><b>${c.Saludable}</b><span>Saludables</span></div><div class=sum><b>${c.Desviado}</b><span>Desviados</span></div><div class=sum><b>${c["Crítico"]}</b><span>Críticos</span></div>`;
  renderEtapas();
  $("tabla").innerHTML = filtrados.length
    ? filtrados
        .map(
          (p, i) =>
            `<tr class=row data-i="${i}"><td>${esc(visible(p.Folio))}</td><td><b>${esc(visible(p.Nombre))}</b></td><td>${esc(visible(convertirJP(p.JP)))}</td><td>${esc(visible(p.LI))}</td><td>${esc(visible(p["Avance Plan"]))}</td><td>${esc(visible(p["Avance real"]))}</td><td>${esc(visible(p["Desviación"]))}</td><td class="health ${p._salud[1]}">${p._salud[0]}</td></tr>`,
        )
        .join("")
    : "<tr><td colspan=8 class=empty>No hay proyectos para mostrar.</td></tr>";
  $("tabla")
    .querySelectorAll(".row")
    .forEach((r) => (r.onclick = () => mostrar(filtrados[+r.dataset.i])));
}
function renderEtapas() {
  let conteo = {};
  ETAPAS.forEach((e) => (conteo[e] = 0));
  filtrados.forEach((p) => {
    let e = visible(p["Etapa"]);
    if (e !== "-") {
      if (conteo[e] === undefined) conteo[e] = 0;
      conteo[e]++;
    }
  });
  let max = Math.max(...Object.values(conteo), 1);
  $("distribucionEtapas").innerHTML = Object.entries(conteo)
    .map(([e, n]) => {
      let ancho = n === 0 ? 0 : Math.max((n / max) * 100, 4);
      return `<div class=stage-row><div class=stage-name>${esc(e)}</div><div class=stage-track><div class=stage-fill style="width:${ancho}%"></div></div><div class=stage-count>${n}</div></div>`;
    })
    .join("");
}
function met(l, v) {
  return `<div class=metric><label>${l}</label><strong>${esc(visible(v))}</strong></div>`;
}
function mostrar(p) {
  proyectoSeleccionado = p;
  $("general").classList.add("hidden");
  $("detalle").classList.remove("hidden");
  let h = p._salud;
  let retraso = visible(p["Días retraso"]);
  let claseRetraso =
    retraso !== "-" && Number(String(retraso).replace(",", ".")) > 0
      ? "delay-positive"
      : "delay-zero";
  $("contenido").innerHTML =
    `<div class=card><div class=detail-head><div><h2>${esc(visible(p.Nombre))} <span class="muted"><b> ${esc(visible(p.Folio))}</b></span></h2><span class=muted>Product Owner: <b>${esc(visible(p.LI))}</b></span><span class=muted style="margin-left: 25px">Jefe Proyecto TI: <b>${esc(visible(convertirJP(p.JP)))}</b></span></div><div style="display: flex; flex-direction: column;"><span class=detail-stage>${esc(visible(p.Etapa))}</span><span class=detail-stage-${h[1]} style="text-align: center;">${esc(visible(h[0]))}</span></div></div></div><div class=card><div class=metrics>${met("Fecha inicio", convertirFecha(p["Fecha inicio"]))}${met("Fecha término", convertirFecha(p["Fecha término"]))}${met("Avance Plan", p["Avance Plan"])}${met("Avance real", p["Avance real"])}${met("Desviación", p["Desviación"])}<div class=metric><label>Días retraso</label><strong class=${claseRetraso}>${esc(retraso)}</strong></div></div><div id=grafico></div></div><div class=card><div class=two><div><h3>Beneficios esperados</h3><div class=info>${esc(visible(p["Beneficios esperados"]))}</div></div><div><h3>Principales Hitos</h3><div class=info>${esc(visible(p["Principales Hitos"]))}</div></div></div></div><div class=card><div class=two><div><h3>Riesgos</h3><div class=risk-box><div class=risk-icon>!</div><div class=risk-content><h4>Riesgos informados</h4><div style="white-space:pre-line">${esc(visible(p.Riesgos))}</div></div></div></div><div><div style="display: flex;gap: 2%;"><h3>Accionables</h3> <span class="detail-stage" style="margin-top: 10px; margin-bottom: 10px; align-content: center;">Responsable: ${esc(visible(p.Responsable))}</span></div><div class=action-box>${esc(visible(p.Accionables))}</div></div></div></div><div class=card><div class=activities><div><h3>Actividad semana en curso</h3><div class=info>${esc(visible(p["Actividades semana en curso"]))}</div></div><div><h3>Actividades semana próxima</h3><div class=info>${esc(visible(p["Actividades semana próxima"]))}</div></div></div></div>`;
  grafico(p);
}
async function generarImagenReporte() {
  const reporte = $("contenido");
  if (!reporte || reporte.classList.contains("hidden"))
    throw new Error("No hay un proyecto seleccionado.");
  return await html2canvas(reporte, {
    backgroundColor: "#f4f6f8",
    scale: 2,
    useCORS: true,
    logging: false,
  });
}

async function descargarReporte() {
  const boton = $("descargarReporte");
  boton.disabled = true;
  try {
    const canvas = await generarImagenReporte();
    const nombre =
      visible(proyectoSeleccionado?.Nombre || "Proyecto")
        .replace(/[^a-z0-9áéíóúñü _-]/gi, "")
        .trim()
        .replace(/\s+/g, "_") || "Proyecto";
    const enlace = document.createElement("a");
    enlace.download = `Reporte_${nombre}.png`;
    enlace.href = canvas.toDataURL("image/png");
    enlace.click();
  } catch (err) {
    alert("No fue posible generar el PNG: " + err.message);
  } finally {
    boton.disabled = false;
  }
}

async function copiarReporte() {
  const boton = $("copiarReporte");
  boton.disabled = true;
  try {
    if (!navigator.clipboard || typeof ClipboardItem === "undefined")
      throw new Error(
        "Tu navegador no permite copiar imágenes directamente al portapapeles.",
      );
    const canvas = await generarImagenReporte();
    const blob = await new Promise((resolve) =>
      canvas.toBlob(resolve, "image/png"),
    );
    if (!blob) throw new Error("No se pudo crear la imagen.");
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    const texto = boton.textContent;
    boton.textContent = "¡Copiado!";
    setTimeout(() => (boton.textContent = texto), 1500);
  } catch (err) {
    alert("No fue posible copiar el reporte: " + err.message);
  } finally {
    boton.disabled = false;
  }
}

function grafico(p) {
  let plan = pct(p["Avance Plan"]),
    real = pct(p["Avance real"]);
  $("grafico").innerHTML =
    plan === null && real === null
      ? "<div class=empty>No hay datos de avance para representar.</div>"
      : [
          ["Plan", plan, "plan"],
          ["Real", real, "real"],
        ]
          .map(
            (x) =>
              `<div class=bar><b>${x[0]}</b><div class=track>${x[1] === null ? "" : `<div class="fill ${x[2]}" style="width:${Math.max(0, Math.min(100, x[1]))}%"></div>`}</div><span>${x[1] === null ? "-" : x[1] + "%"}</span></div>`,
          )
          .join("");
}

$("volver").onclick = () => {
  $("detalle").classList.add("hidden");
  $("general").classList.remove("hidden");
};
$("descargarReporte").onclick = descargarReporte;
$("copiarReporte").onclick = copiarReporte;
$("tabla").innerHTML =
  "<tr><td colspan=8 class=empty>Selecciona el Excel maestro para cargar los proyectos.</td></tr>";

async function crearCorreoOutlook() {
  if (!proyectoSeleccionado) {
    alert("No hay un proyecto seleccionado.");
    return;
  }

  const canvas = await copiarReporte();

  const nombre = visible(proyectoSeleccionado["Nombre"]);

  const asunto = `Avance semanal - ${nombre}`;

  const cuerpo = `Estimados,

Comparto el avance semanal del proyecto ${nombre}.

[PEGAR IMAGEN DEL REPORTE AQUÍ]

Saludos.`;

  const mailto =
    `mailto:?subject=${encodeURIComponent(asunto)}` +
    `&body=${encodeURIComponent(cuerpo)}`;

  window.location.href = mailto;
}
$("crearCorreo").addEventListener("click", crearCorreoOutlook);

function convertirFecha(fecha) {
  if (fecha === "TBD") {
    return "Por Definir";
  }else if (fecha != null && fecha != undefined && fecha.trim() !== "") {
    const [mm, dd, yyyy] = fecha.split("/");
    return `${dd.padStart(2, "0")}/${mm.padStart(2, "0")}/${yyyy.length === 2 ? "20" + yyyy : yyyy}`;
  } else {
    return "-";
  }
}

function convertirJP(jp) {
  jp_completo = jp;
  switch (jp) {
    case (jp = "D_Vidal"):
      return "Daniel Vidal";
    case (jp = "J_Vilches"):
      return "Jordan Vilches";
    case (jp = "J_Meza"):
      return "Jaime Meza";
    case (jp = "O_Moris"):
      return "Omer Moris";
    case (jp = "P_Reyes"):
      return "Pamela Reyes";
    default:
      return jp_completo;
      break;
  }
}
