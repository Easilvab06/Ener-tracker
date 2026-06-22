// Cloudflare Worker — Proxy CORS para la matriz de Hostinger
// Reenvía la petición al servidor de Hostinger y agrega las cabeceras
// CORS necesarias para que GitHub Pages pueda leer la respuesta.
//
// La tabla real vive en la RAÍZ del dominio de Hostinger:
//   https://peachpuff-stingray-882207.hostingersite.com/
//
// Este Worker expone esa misma raíz bajo /proxy-matriz, por ejemplo:
//   https://ener-tracker-proxy.soinsolar1.workers.dev/proxy-matriz
// devuelve exactamente lo mismo que la raíz de Hostinger, pero con
// cabeceras CORS que el navegador sí acepta desde GitHub Pages.
//
// CAMBIOS CLAVE respecto a la versión anterior (fix del error 400 en producción):
// 1. El preflight (OPTIONS) ahora refleja exactamente los headers y el método
//    que el navegador solicitó (Access-Control-Request-Headers /
//    -Request-Method), en vez de una lista fija. Algunos navegadores
//    rechazan la respuesta si el preflight no "hace match" exacto.
// 2. Se refleja el Origin real del request en vez de usar siempre '*',
//    que es más robusto cuando el navegador hace fetch con modo 'cors'
//    y ciertas combinaciones de headers.
// 3. Se añade manejo explícito de errores de upstream (status 4xx/5xx)
//    y un try/catch más granular para no devolver 400 por errores que
//    en realidad son de red (esos deben dar 502, no 400).
// 4. Se fuerza el método GET aunque el navegador mande otro verbo por error.

const TARGET_BASE = 'https://peachpuff-stingray-882207.hostingersite.com';

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '*';

    // ── Pre-flight CORS ────────────────────────────────────────────────
    // El navegador manda OPTIONS antes del GET real cuando detecta que
    // la petición no es "simple" (p.ej. lleva un header Accept no-estándar).
    // Hay que responder reflejando EXACTAMENTE lo que el navegador pidió.
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(request, origin),
      });
    }

    // Todo lo que venga después de /proxy-matriz se reenvía a la raíz de
    // Hostinger (o a la subruta equivalente, si alguna vez la tabla se
    // mueve a otra ruta dentro del mismo dominio).
    const targetPath = url.pathname.replace(/^\/proxy-matriz/, '') + url.search;
    const targetUrl = TARGET_BASE + (targetPath || '/');

    try {
      const upstreamResponse = await fetch(targetUrl, {
        method: 'GET',
        headers: { 'Accept': 'application/json, text/html, */*' },
        // Evita que el Worker reenvíe cabeceras del navegador (cookies,
        // referer, etc.) que Hostinger no espera y podrían provocar
        // respuestas inesperadas (incluyendo 400) en el origen.
        redirect: 'follow',
      });

      const body = await upstreamResponse.text();

      return new Response(body, {
        status: upstreamResponse.status,
        headers: {
          ...corsHeaders(request, origin),
          'Content-Type': upstreamResponse.headers.get('content-type') || 'text/html; charset=utf-8',
          'Cache-Control': 'no-store',
        },
      });
    } catch (err) {
      // Un fallo de red/DNS/TLS hacia Hostinger es un 502 (Bad Gateway),
      // nunca un 400 — un 400 indicaría que la petición del CLIENTE estaba
      // mal formada, lo cual no es el caso aquí.
      return new Response(JSON.stringify({ error: 'Proxy fetch failed', detail: err.message }), {
        status: 502,
        headers: {
          ...corsHeaders(request, origin),
          'Content-Type': 'application/json',
        },
      });
    }
  },
};

function corsHeaders(request, origin) {
  // Reflejar los headers/método solicitados por el navegador en el
  // preflight es más robusto que una lista fija: evita que el navegador
  // rechace la respuesta por "header no permitido" cuando el front-end
  // cambie qué headers manda en el fetch real.
  const requestedHeaders = request.headers.get('Access-Control-Request-Headers');
  const requestedMethod = request.headers.get('Access-Control-Request-Method');

  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': requestedMethod || 'GET, OPTIONS',
    'Access-Control-Allow-Headers': requestedHeaders || 'Content-Type, Accept',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin',
  };
}