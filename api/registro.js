// Registro de uso del tablero: quien entra y que informe abre.
//
// El tablero le avisa (navigator.sendBeacon) cada vez que alguien entra y cada
// vez que hace clic en una tarjeta. Aca se guarda una linea por evento en
// Upstash Redis, una lista por dia ("reg:2026-10-29"), que vence sola al año.
// El orquestador la baja con el token de SOLO LECTURA (tablero_registro.py).
//
// Solo escribe: no devuelve nada de lo guardado. El usuario lo informa el
// navegador (el login del tablero es estatico), asi que sirve para medir uso,
// no como auditoria.

const EVENTOS = new Set(["login", "sesion", "clic"]);
const VENCE_SEG = 400 * 86400;

const corto = (v, n) => (v ?? "").toString().slice(0, n);

function hoyArgentina() {
  // en-CA da YYYY-MM-DD
  return new Date().toLocaleDateString("en-CA",
    { timeZone: "America/Argentina/Buenos_Aires" });
}

module.exports = async (req, res) => {
  if (req.method !== "POST") { res.status(405).end(); return; }

  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) { res.status(503).end(); return; }

  let b = req.body;
  if (typeof b === "string") { try { b = JSON.parse(b); } catch (e) { b = null; } }
  if (!b || !EVENTOS.has(b.evento)) { res.status(400).end(); return; }

  const linea = JSON.stringify({
    ts: new Date().toISOString(),
    usuario: corto(b.usuario, 60).toLowerCase(),
    evento: b.evento,
    informe: corto(b.informe, 200),
    mes: corto(b.mes, 60),
    area: corto(b.area, 80),
    url: corto(b.url, 500),
    movil: /Mobi|Android|iPhone/i.test(req.headers["user-agent"] || ""),
  });

  const clave = "reg:" + hoyArgentina();
  try {
    const r = await fetch(url.replace(/\/$/, "") + "/pipeline", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify([["RPUSH", clave, linea], ["EXPIRE", clave, String(VENCE_SEG)]]),
    });
    res.status(r.ok ? 204 : 502).end();
  } catch (e) {
    res.status(502).end();
  }
};
