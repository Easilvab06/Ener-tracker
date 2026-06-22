// ═══════════════════════════════════════════════════════════════════
//  useSolarData.js  —  Composable adaptado a matriz_completa.xlsx
//  + Soporte para fuentes HTML (tabla dinámica web) y JSON
//
//  Columnas de la nueva matriz:
//    indice | Timestamp | Temp panel | Temp ambiente | Irradiancia
//    Vi1 Vi2 Vi3 | Ci1 Ci2 Ci3 | Pi1 Pi2 Pi3
//    Vc1 Vc2 Vc3 | Cc1 Cc2 Cc3 | Pc1 Pc2 Pc3
// ═══════════════════════════════════════════════════════════════════

import { ref, computed, watch } from 'vue'
import * as XLSX from 'xlsx'

// ─── helpers ────────────────────────────────────────────────────────
const avg  = arr => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0
const sum  = arr => arr.reduce((s, v) => s + v, 0)
const safe = v   => (v === null || v === undefined || isNaN(v)) ? 0 : +v

/**
 * Parser de timestamp EXPLÍCITO — evita new Date(string), que es ambiguo
 * entre navegadores y formatos (confunde DD/MM con MM/DD, falla con
 * formatos no-ISO, etc.). Reconoce los formatos que la tabla de Hostinger
 * suele emitir:
 *   - "DD/MM/YYYY, HH:mm:ss"  o  "DD/MM/YYYY HH:mm:ss"
 *   - "YYYY-MM-DD HH:mm:ss"   o  "YYYY-MM-DDTHH:mm:ss" (ISO)
 *   - "DD-MM-YYYY HH:mm:ss"
 * Devuelve milisegundos (timestamp local) o NaN si no pudo parsear.
 */
function parseTimestampString(str) {
  const s = str.trim()

  // Formato ISO: YYYY-MM-DD[ T]HH:mm:ss
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/)
  if (m) {
    const [, y, mo, d, h, mi, se] = m.map(Number)
    return new Date(y, mo - 1, d, h, mi, se).getTime()
  }

  // Formato DD/MM/YYYY o DD-MM-YYYY, con hora opcional, coma opcional
  m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/)
  if (m) {
    const [, d, mo, y, h, mi, se] = m.map(v => v === undefined ? 0 : Number(v))
    return new Date(y, mo - 1, d, h, mi, se || 0).getTime()
  }

  // Solo fecha sin hora: DD/MM/YYYY o YYYY-MM-DD
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (m) {
    const [, y, mo, d] = m.map(Number)
    return new Date(y, mo - 1, d).getTime()
  }
  m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/)
  if (m) {
    const [, d, mo, y] = m.map(Number)
    return new Date(y, mo - 1, d).getTime()
  }

  // Último recurso: dejar que el navegador intente (puede fallar/ser ambiguo)
  const fallback = new Date(s).getTime()
  return isNaN(fallback) ? NaN : fallback
}

/** Convierte una fila del Excel (objeto con las claves originales)
 *  al formato interno normalizado que usa la app. */
function normalizeRow(raw) {
  let ts
  const rawTs = raw['Timestamp'] ?? raw['timestamp'] ?? raw['ts'] ?? raw['tsMs']
  if (typeof rawTs === 'number') {
    const d = XLSX.SSF.parse_date_code(rawTs)
    ts = new Date(d.y, d.m - 1, d.d, d.H, d.M, d.S).getTime()
  } else if (rawTs instanceof Date) {
    ts = rawTs.getTime()
  } else if (typeof rawTs === 'string') {
    ts = parseTimestampString(rawTs)
    if (isNaN(ts)) ts = Date.now() // fallback si de verdad no se pudo parsear
  } else {
    ts = Date.now()
  }

  const vi1 = safe(raw['Vi1'] ?? raw['vi1'])
  const vi2 = safe(raw['Vi2'] ?? raw['vi2'])
  const vi3 = safe(raw['Vi3'] ?? raw['vi3'])

  const ci1 = safe(raw['Ci1'] ?? raw['ci1'])
  const ci2 = safe(raw['Ci2'] ?? raw['ci2'])
  const ci3 = safe(raw['Ci3'] ?? raw['ci3'])

  const pi1 = safe(raw['Pi1'] ?? raw['pi1'])
  const pi2 = safe(raw['Pi2'] ?? raw['pi2'])
  const pi3 = safe(raw['Pi3'] ?? raw['pi3'])

  const vc1 = safe(raw['Vc1'] ?? raw['vc1'])
  const vc2 = safe(raw['Vc2'] ?? raw['vc2'])
  const vc3 = safe(raw['Vc3'] ?? raw['vc3'])

  const cc1 = safe(raw['Cc1'] ?? raw['cc1'])
  const cc2 = safe(raw['Cc2'] ?? raw['cc2'])
  const cc3 = safe(raw['Cc3'] ?? raw['cc3'])

  const pc1 = safe(raw['Pc1'] ?? raw['pc1'])
  const pc2 = safe(raw['Pc2'] ?? raw['pc2'])
  const pc3 = safe(raw['Pc3'] ?? raw['pc3'])

  return {
    tsMs:      ts,
    tempPanel: safe(raw['Temp panel']    ?? raw['tempPanel'] ?? raw['temp_panel']),
    tempAmb:   safe(raw['Temp ambiente'] ?? raw['tempAmb']   ?? raw['temp_ambiente']),
    irrad:     safe(raw['Irradiancia']   ?? raw['irrad']     ?? raw['irradiancia']),

    vi1, vi2, vi3,
    ci1, ci2, ci3,
    pi1, pi2, pi3,

    vc1, vc2, vc3,
    cc1, cc2, cc3,
    pc1, pc2, pc3,

    pti: pi1 + pi2 + pi3,
    ptc: pc1 + pc2 + pc3,
    vi:  avg([vi1, vi2, vi3].filter(v => v > 0)) || 0,
    ci:  avg([ci1, ci2, ci3].filter(v => v > 0)) || 0,
    vc:  avg([vc1, vc2, vc3].filter(v => v > 0)) || 0,
    cc:  avg([cc1, cc2, cc3].filter(v => v > 0)) || 0,
  }
}

// ─── HTML table parser ────────────────────────────────────────────────
/**
 * Parsea una respuesta HTML que contiene una <table> con la misma
 * estructura de columnas que la matriz Excel.
 * Devuelve un array de objetos crudos listos para normalizeRow().
 */
function parseHtmlTable(html) {
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')
  const table = doc.querySelector('table')
  if (!table) throw new Error('No se encontró ninguna tabla <table> en la respuesta HTML.')

  const rows = Array.from(table.querySelectorAll('tr'))
  if (rows.length < 2) throw new Error('La tabla no tiene datos suficientes.')

  // Extraer cabeceras de la primera fila (th o td)
  const headers = Array.from(rows[0].querySelectorAll('th, td')).map(cell =>
    cell.textContent.trim()
  )

  // Convertir cada fila de datos en un objeto { header: value }
  const result = []
  for (let i = 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll('td'))
    if (cells.length === 0) continue
    const obj = {}
    headers.forEach((h, idx) => {
      const raw = cells[idx]?.textContent.trim() ?? ''
      // No tocar la columna Timestamp: parseFloat() trunca "17/06/2026..." al
      // número 17, perdiendo toda la fecha. Para todo lo demás, intentar número.
      const isTimestampCol = /^timestamp$|^ts$|^fecha$/i.test(h)
      if (isTimestampCol) {
        obj[h] = raw
      } else {
        const num = parseFloat(raw)
        obj[h] = (raw !== '' && !isNaN(num)) ? num : raw
      }
    })
    result.push(obj)
  }
  return result
}

// ─── constants ──────────────────────────────────────────────────────
const CO2_KG_PER_KWH  = 0.233
const TREE_KG_PER_YR  = 22

// ─── resolución de URL del proxy según entorno ─────────────────────
// En desarrollo (npm run dev) Vite intercepta '/proxy-matriz' y lo
// reenvía desde su propio servidor de Vite — sin problema de CORS.
//
// En producción (GitHub Pages) no hay servidor Vite, así que el navegador
// pega DIRECTO a Hostinger y el navegador bloquea la respuesta por CORS
// (Hostinger no manda la cabecera Access-Control-Allow-Origin).
//
// IMPORTANTE (descubierto en producción): los proxies CORS públicos
// gratuitos tienen límites de tamaño de respuesta muy bajos
// (corsproxy.io: 1MB → 413, codetabs: límite similar → 400) y la tabla
// HTML de Hostinger ya los supera a medida que crece con el tiempo.
// allorigins.win además ha estado intermitente/caído. Por eso estos
// proxies YA NO SON CONFIABLES como fuente principal — se prueba primero,
// y casi exclusivamente, el Worker propio de Cloudflare, que no tiene
// ese límite de tamaño. Los proxies públicos quedan solo como respaldo
// de último recurso, sabiendo que probablemente fallarán si la tabla es
// grande.
const HOSTINGER_BASE = 'https://peachpuff-stingray-882207.hostingersite.com'

// Worker propio (Cloudflare) — sin límite de tamaño de payload práctico
// (el plan gratis soporta respuestas de hasta 100MB+), a diferencia de los
// proxies públicos gratuitos. Es la fuente principal y debería ser
// suficiente por sí sola si está bien desplegado.
const OWN_WORKER_PROXY = 'https://ener-tracker-proxy.soinsolar1.workers.dev/proxy-matriz'

const PUBLIC_CORS_PROXIES = [
  // 1. Worker propio: reenvía directo a la raíz de Hostinger, sin tocar la URL,
  //    y sin el límite de tamaño que tienen los proxies públicos gratuitos.
  { build: () => OWN_WORKER_PROXY, name: 'Worker propio' },
  // 2-3. Proxies públicos de respaldo — casi seguro fallarán si la tabla supera
  //    ~1MB, pero se dejan por si Hostinger reduce el tamaño de la tabla
  //    (p.ej. paginación) o el Worker propio estuviera caído.
  { build: url => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`, name: 'allorigins' },
  { build: url => `https://api.codetabs.com/v1/proxy/?quest=${encodeURIComponent(url)}`, name: 'codetabs' },
  // corsproxy.io al final: limita su tier gratis a 1MB de payload y, desde
  // 2026, a peticiones desde localhost — casi nunca funcionará en producción.
  { build: url => `https://corsproxy.io/?url=${encodeURIComponent(url)}`, name: 'corsproxy.io' },
]

function toRealUrl(url) {
  // Caso 1: ruta interna '/proxy-matriz' (la usada en desarrollo) → raíz real
  if (url.startsWith('/proxy-matriz')) return HOSTINGER_BASE + url.replace('/proxy-matriz', '')

  // Caso 2: alguien pegó manualmente la URL completa de Hostinger con
  // '/proxy-matriz' al final (ruta que NO existe en el servidor real,
  // la tabla vive en la raíz). La normalizamos también a la raíz.
  if (url.startsWith(HOSTINGER_BASE + '/proxy-matriz')) {
    return url.replace(HOSTINGER_BASE + '/proxy-matriz', HOSTINGER_BASE)
  }

  return url
}

/** ¿Esta URL necesita pasar por un proxy CORS en producción? */
function needsProxy(url) {
  if (import.meta.env.DEV) return false // Vite ya lo maneja en local
  return url.startsWith('/proxy-matriz') || url.startsWith(HOSTINGER_BASE)
}

/**
 * Hace fetch probando, en orden, los proxies CORS disponibles, empezando
 * SIEMPRE por el Worker propio (sin límite de tamaño de payload).
 * Si todos fallan, lanza un error con el detalle de CADA proxy y su causa
 * específica (código HTTP o tipo de fallo), para poder diagnosticar sin
 * tener que abrir DevTools.
 *
 * Si la URL no necesita proxy (desarrollo, o ya es una API externa con
 * CORS habilitado), hace el fetch directo normal.
 *
 * IMPORTANTE: esta versión NO manda headers personalizados (como un
 * 'Accept' no-estándar) en la petición al proxy. Un header de ese tipo
 * convierte el GET en una petición "no simple" y obliga al navegador a
 * hacer un preflight OPTIONS antes del GET real — si el proxy de turno no
 * maneja ese preflight exactamente como el navegador espera, la petición
 * falla con un error de red que en la consola se ve como CORS/400, aunque
 * el proxy "funcione" cuando se prueba desde curl o un servidor (que no
 * hacen preflight). Quitar el header evita ese preflight y hace el fetch
 * mucho más confiable en producción.
 *
 * Cada intento de proxy tiene su propio timeout corto (6s) para que, si el
 * caller pasó un AbortSignal de timeout más largo, igual se pueda pasar al
 * siguiente proxy de la lista en vez de agotar todo el tiempo en el primero.
 */
export async function fetchSourceWithFallback(url, fetchOpts = {}) {
  if (!needsProxy(url)) {
    return fetch(url, fetchOpts)
  }

  // Importante: NO reutilizamos fetchOpts.signal para cada intento individual,
  // y tampoco reutilizamos fetchOpts.headers — ver nota arriba sobre por qué
  // evitar headers personalizados en la petición al proxy.
  const { signal: _ignoredSignal, headers: _ignoredHeaders, ...restOpts } = fetchOpts
  const realUrl = toRealUrl(url)
  const attempts = []

  for (const proxy of PUBLIC_CORS_PROXIES) {
    try {
      const res = await fetch(proxy.build(realUrl), {
        ...restOpts,
        signal: AbortSignal.timeout(8000),
      })
      if (!res.ok) {
        const reason = res.status === 413
          ? 'respuesta demasiado grande para este proxy'
          : `HTTP ${res.status}`
        attempts.push(`${proxy.name}: ${reason}`)
        continue
      }
      return res
    } catch (err) {
      const isTimeout = err.name === 'TimeoutError' || err.name === 'AbortError'
      const isCors = err instanceof TypeError // fetch lanza TypeError genérico en fallos de CORS/red
      attempts.push(`${proxy.name}: ${isTimeout ? 'tiempo de espera agotado' : isCors ? 'bloqueado por CORS o proxy caído' : err.message}`)
      // Probar el siguiente proxy de la lista
    }
  }

  const workerFailed = attempts[0]?.startsWith('Worker propio')
  const hint = workerFailed
    ? ' — revisa que el Worker esté desplegado (wrangler deploy) y accesible.'
    : ''

  throw new Error(`No se pudo conectar a través de ningún proxy disponible.${hint} Detalle: ${attempts.join(' | ')}`)
}

// Se mantiene exportada por compatibilidad, pero ya no se usa para construir
// la URL final — ahora fetchSourceWithFallback maneja todo el flujo con fallback.
export function resolveSourceUrl(url) {
  return url
}

// ─── composable ─────────────────────────────────────────────────────
export function useSolarData() {
  const allRows    = ref([])
  const fileName   = ref('')
  const isLoading  = ref(false)
  const loadProgress = ref(0)
  const isComputing  = ref(false)
  const lastUpdated  = ref(null)

  const timeRange  = ref('day')
  const activeDate = ref(new Date())
  const activeInv  = ref('all')

  const autoRefresh = ref(false)
  const refreshSec  = ref(30)
  let   _timer      = null

  const apiMode      = ref(false)
  const apiUrl       = ref('')
  const apiConnected = ref(false)
  const apiError     = ref('')
  const liveTs       = ref(null)

  // ─── Estadísticas de la fuente para la pantalla de conexión ────────
  // Expone info para que UploadScreen pueda mostrar un preview
  const apiSourceInfo = ref(null)  // { total, first, last }

  // ── filtrado temporal ──────────────────────────────────────────────
  const filteredRows = computed(() => {
    if (!allRows.value.length) return []
    const d = activeDate.value
    const y = d.getFullYear()
    const m = d.getMonth()
    const day = d.getDate()

    return allRows.value.filter(r => {
      const rd = new Date(r.tsMs)
      if (timeRange.value === 'day')   return rd.getFullYear()===y && rd.getMonth()===m && rd.getDate()===day
      if (timeRange.value === 'month') return rd.getFullYear()===y && rd.getMonth()===m
      return rd.getFullYear() === y
    })
  })

  const liveRecord = computed(() => {
    const rows = filteredRows.value
    return rows.length ? rows[rows.length - 1] : null
  })

  const availableDays = computed(() => {
    const s = new Set()
    allRows.value.forEach(r => {
      const d = new Date(r.tsMs)
      const y = d.getFullYear()
      const m = String(d.getMonth()+1).padStart(2,'0')
      const dd= String(d.getDate()).padStart(2,'0')
      s.add(`${y}-${m}-${dd}`)
    })
    return [...s].sort()
  })

  const availableMonths = computed(() => {
    const s = new Set()
    allRows.value.forEach(r => {
      const d = new Date(r.tsMs)
      const y = d.getFullYear()
      const m = String(d.getMonth()+1).padStart(2,'0')
      s.add(`${y}-${m}`)
    })
    return [...s].sort()
  })

  const availableYears = computed(() => {
    const s = new Set(allRows.value.map(r => String(new Date(r.tsMs).getFullYear())))
    return { value: [...s].sort() }
  })

  // ── KPIs ───────────────────────────────────────────────────────────
  const kpis = computed(() => {
    const rows = filteredRows.value
    if (!rows.length) return null

    const ptiVals = rows.map(r => r.pti)
    const ptcVals = rows.map(r => r.ptc)

    const totalPti = sum(ptiVals)
    const totalPtc = sum(ptcVals)
    const maxPower = Math.max(...ptiVals)

    const isDayView = timeRange.value === 'day'

    let ptiDisp, ptcDisp, maxDisp, ptiUnit, ptcUnit
    if (isDayView) {
      ptiDisp = (avg(ptiVals)).toFixed(1)
      ptcDisp = (avg(ptcVals)).toFixed(1)
      maxDisp = maxPower.toFixed(1)
      ptiUnit = 'W'; ptcUnit = 'W'
    } else {
      const meanIntervalH = rows.length > 1
        ? (rows[rows.length-1].tsMs - rows[0].tsMs) / 3600000 / (rows.length - 1)
        : 1/60
      const kwhPti = totalPti * meanIntervalH / 1000
      const kwhPtc = totalPtc * meanIntervalH / 1000
      const kwhMax = maxPower * meanIntervalH / 1000
      ptiDisp = kwhPti >= 1000 ? (kwhPti/1000).toFixed(2) : kwhPti.toFixed(1)
      ptcDisp = kwhPtc >= 1000 ? (kwhPtc/1000).toFixed(2) : kwhPtc.toFixed(1)
      maxDisp = kwhMax >= 1000 ? (kwhMax/1000).toFixed(2) : kwhMax.toFixed(1)
      ptiUnit = kwhPti >= 1000 ? 'MWh' : 'kWh'
      ptcUnit = kwhPtc >= 1000 ? 'MWh' : 'kWh'
    }

    const ptiTotal = sum(ptiVals)
    const ptcTotal = sum(ptcVals)
    const eff = ptiTotal > 0 ? ((ptcTotal / ptiTotal) * 100).toFixed(1) : '0.0'

    const irrad     = avg(rows.map(r => r.irrad)).toFixed(0)
    const tempPanel = avg(rows.map(r => r.tempPanel)).toFixed(1)
    const tempAmb   = avg(rows.map(r => r.tempAmb)).toFixed(1)

    const viVals = rows.flatMap(r => [r.vi1, r.vi2, r.vi3].filter(v => v > 0))
    const ciVals = rows.flatMap(r => [r.ci1, r.ci2, r.ci3].filter(v => v > 0))
    const vi = viVals.length ? avg(viVals).toFixed(1) : '0.0'
    const ci = ciVals.length ? avg(ciVals).toFixed(2) : '0.00'

    const activeRows = rows.filter(r => r.irrad > 100)
    const peakHours = rows.length > 1
      ? (activeRows.length * (rows[rows.length-1].tsMs - rows[0].tsMs) / (rows.length - 1) / 3600000).toFixed(1)
      : '0.0'

    const kwhGen = isDayView
      ? (avg(ptiVals) * (rows.length > 1 ? (rows[rows.length-1].tsMs - rows[0].tsMs)/3600000/(rows.length-1) : 1/60) * rows.length / 1000)
      : (ptiTotal * (rows.length > 1 ? (rows[rows.length-1].tsMs - rows[0].tsMs)/3600000/(rows.length-1) : 1/60) / 1000)
    const co2   = (kwhGen * CO2_KG_PER_KWH / 1000).toFixed(3)
    const trees = Math.round(kwhGen * CO2_KG_PER_KWH / TREE_KG_PER_YR)

    return { pti: ptiDisp, ptiUnit, ptc: ptcDisp, ptcUnit, maxPower: maxDisp, eff, irrad, tempPanel, tempAmb, vi, ci, peakHours, co2, trees }
  })

  // ── Datos para gráficos ────────────────────────────────────────────
  const chartData = computed(() => {
    const rows = filteredRows.value
    if (!rows.length) return null

    const isDayView = timeRange.value === 'day'
    let labels, grouped

    if (isDayView) {
      labels  = rows.map(r => {
        const d = new Date(r.tsMs)
        return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`
      })
      grouped = rows
    } else if (timeRange.value === 'month') {
      const map = {}
      rows.forEach(r => {
        const d   = new Date(r.tsMs)
        const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
        ;(map[key] = map[key] || []).push(r)
      })
      labels  = Object.keys(map).sort()
      grouped = labels.map(k => avgGroup(map[k]))
    } else {
      const map = {}
      rows.forEach(r => {
        const d   = new Date(r.tsMs)
        const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`
        ;(map[key] = map[key] || []).push(r)
      })
      labels  = Object.keys(map).sort()
      grouped = labels.map(k => avgGroup(map[k]))
    }

    return {
      isDayView,
      labels,
      pti:       grouped.map(r => r.pti),
      ptc:       grouped.map(r => r.ptc),
      irrad:     grouped.map(r => r.irrad),
      tempPanel: grouped.map(r => r.tempPanel),
      tempAmb:   grouped.map(r => r.tempAmb),
      vi:        grouped.map(r => r.vi),
      ci:        grouped.map(r => r.ci),
      vi1: grouped.map(r => r.vi1), vi2: grouped.map(r => r.vi2), vi3: grouped.map(r => r.vi3),
      ci1: grouped.map(r => r.ci1), ci2: grouped.map(r => r.ci2), ci3: grouped.map(r => r.ci3),
      vc1: grouped.map(r => r.vc1), vc2: grouped.map(r => r.vc2), vc3: grouped.map(r => r.vc3),
      cc1: grouped.map(r => r.cc1), cc2: grouped.map(r => r.cc2), cc3: grouped.map(r => r.cc3),
      inv1_pi: grouped.map(r => r.pi1), inv2_pi: grouped.map(r => r.pi2), inv3_pi: grouped.map(r => r.pi3),
      inv1_pc: grouped.map(r => r.pc1), inv2_pc: grouped.map(r => r.pc2), inv3_pc: grouped.map(r => r.pc3),
    }
  })

  function avgGroup(rows) {
    const n = rows.length || 1
    const s = k => rows.reduce((a, r) => a + (r[k] || 0), 0) / n
    return {
      pti: s('pti'), ptc: s('ptc'), irrad: s('irrad'),
      tempPanel: s('tempPanel'), tempAmb: s('tempAmb'),
      vi: s('vi'), ci: s('ci'), vc: s('vc'), cc: s('cc'),
      vi1: s('vi1'), vi2: s('vi2'), vi3: s('vi3'),
      ci1: s('ci1'), ci2: s('ci2'), ci3: s('ci3'),
      vc1: s('vc1'), vc2: s('vc2'), vc3: s('vc3'),
      cc1: s('cc1'), cc2: s('cc2'), cc3: s('cc3'),
      pi1: s('pi1'), pi2: s('pi2'), pi3: s('pi3'),
      pc1: s('pc1'), pc2: s('pc2'), pc3: s('pc3'),
    }
  }

  // ── Navegación de período ──────────────────────────────────────────
  function prevPeriod() {
    const d = new Date(activeDate.value)
    if (timeRange.value === 'day')   d.setDate(d.getDate() - 1)
    if (timeRange.value === 'month') d.setMonth(d.getMonth() - 1)
    if (timeRange.value === 'year')  d.setFullYear(d.getFullYear() - 1)
    activeDate.value = d
  }

  function nextPeriod() {
    const d = new Date(activeDate.value)
    if (timeRange.value === 'day')   d.setDate(d.getDate() + 1)
    if (timeRange.value === 'month') d.setMonth(d.getMonth() + 1)
    if (timeRange.value === 'year')  d.setFullYear(d.getFullYear() + 1)
    activeDate.value = d
  }

  const MN = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
  function periodLabel() {
    const d = activeDate.value
    if (timeRange.value === 'day')   return d.toLocaleDateString('es-CO',{day:'2-digit',month:'short',year:'numeric'})
    if (timeRange.value === 'month') return `${MN[d.getMonth()]} ${d.getFullYear()}`
    return String(d.getFullYear())
  }

  // ── Carga de archivo XLSX ──────────────────────────────────────────
  async function loadFile(file) {
    isLoading.value  = true
    loadProgress.value = 0
    apiMode.value    = false
    apiConnected.value = false

    try {
      fileName.value = file.name
      loadProgress.value = 10

      const buffer = await file.arrayBuffer()
      loadProgress.value = 30

      const wb = XLSX.read(buffer, { type: 'array', cellDates: true })
      loadProgress.value = 50

      const sheetName = wb.SheetNames.includes('Matriz')
        ? 'Matriz'
        : wb.SheetNames[0]
      const ws = wb.Sheets[sheetName]

      loadProgress.value = 60

      const rawRows = XLSX.utils.sheet_to_json(ws, { defval: null })
      loadProgress.value = 75

      const CHUNK = 500
      const normalized = []
      for (let i = 0; i < rawRows.length; i += CHUNK) {
        const chunk = rawRows.slice(i, i + CHUNK).map(normalizeRow)
        normalized.push(...chunk)
        loadProgress.value = 75 + Math.round(((i + CHUNK) / rawRows.length) * 20)
        await new Promise(r => setTimeout(r, 0))
      }

      normalized.sort((a, b) => a.tsMs - b.tsMs)
      allRows.value = normalized
      loadProgress.value = 100

      if (normalized.length) {
        activeDate.value = new Date(normalized[normalized.length - 1].tsMs)
      }

      lastUpdated.value = new Date()
    } catch (err) {
      console.error('[useSolarData] Error cargando archivo:', err)
    } finally {
      setTimeout(() => { isLoading.value = false; loadProgress.value = 0 }, 300)
    }
  }

  // ── Conexión API/HTML — detección automática de formato ─────────────
  async function loadFromApi(url) {
    apiMode.value  = true
    apiUrl.value   = url
    apiError.value = ''
    isLoading.value = true
    loadProgress.value = 20

    try {
      const res = await fetchSourceWithFallback(url, {
        headers: { 'Accept': 'application/json, text/html, */*' }
      })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)

      const contentType = res.headers.get('content-type') || ''
      let rawRows = []

      if (contentType.includes('application/json')) {
        // ── Respuesta JSON ─────────────────────────────────────────
        const data = await res.json()
        loadProgress.value = 60
        const rows = Array.isArray(data) ? data : (data.data ?? data.rows ?? [])
        rawRows = rows

      } else {
        // ── Respuesta HTML — parsear la <table> ────────────────────
        const html = await res.text()
        loadProgress.value = 50
        rawRows = parseHtmlTable(html)
        loadProgress.value = 65
      }

      // Normalizar y ordenar
      const normalized = rawRows.map(normalizeRow).sort((a, b) => a.tsMs - b.tsMs)
      allRows.value    = normalized
      apiConnected.value = true
      lastUpdated.value  = new Date()
      liveTs.value       = Date.now()
      fileName.value     = url

      // Info para el panel de estado
      if (normalized.length) {
        const first = new Date(normalized[0].tsMs)
        const last  = new Date(normalized[normalized.length - 1].tsMs)
        apiSourceInfo.value = {
          total: normalized.length,
          first: first.toISOString().substring(0, 10),
          last:  last.toISOString().substring(0, 10),
        }
        activeDate.value = last
      }

      loadProgress.value = 100
    } catch (err) {
      apiError.value     = err.message
      apiConnected.value = false
      console.error('[useSolarData] Error fuente:', err)
    } finally {
      setTimeout(() => { isLoading.value = false; loadProgress.value = 0 }, 300)
    }
  }

  // ── Función de prueba (sin cargar datos, solo estadísticas) ────────
  async function testApiSource(url) {
    try {
      const res = await fetchSourceWithFallback(url, {
        headers: { 'Accept': 'application/json, text/html, */*' },
        signal: AbortSignal.timeout(8000)
      })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)

      const contentType = res.headers.get('content-type') || ''
      let rawRows = []

      if (contentType.includes('application/json')) {
        const data = await res.json()
        const rows = Array.isArray(data) ? data : (data.data ?? data.rows ?? [])
        rawRows = rows
      } else {
        const html = await res.text()
        rawRows = parseHtmlTable(html)
      }

      const normalized = rawRows.map(normalizeRow).sort((a, b) => a.tsMs - b.tsMs)
      if (!normalized.length) throw new Error('La fuente no devolvió datos.')

      return {
        ok: true,
        total: normalized.length,
        first: new Date(normalized[0].tsMs).toISOString().substring(0, 10),
        last:  new Date(normalized[normalized.length - 1].tsMs).toISOString().substring(0, 10),
        type: contentType.includes('json') ? 'JSON' : 'HTML',
      }
    } catch (err) {
      return { ok: false, error: err.message }
    }
  }

  // ── Auto-refresh ───────────────────────────────────────────────────
  function toggleRefresh() {
    autoRefresh.value = !autoRefresh.value
    if (autoRefresh.value) scheduleRefresh()
    else stopRefresh()
  }

  function scheduleRefresh() {
    clearInterval(_timer)
    _timer = setInterval(async () => {
      if (!apiMode.value || !apiUrl.value) return
      isComputing.value = true
      try { await loadFromApi(apiUrl.value) } finally { isComputing.value = false }
    }, refreshSec.value * 1000)
  }

  function stopRefresh() {
    clearInterval(_timer)
    autoRefresh.value = false
  }

  watch(refreshSec, () => { if (autoRefresh.value) scheduleRefresh() })

  // ── Exportar ───────────────────────────────────────────────────────
  return {
    // Estado
    allRows, fileName, isLoading, loadProgress, isComputing, lastUpdated,
    // Controles
    timeRange, activeDate, activeInv,
    autoRefresh, refreshSec,
    // API
    apiMode, apiUrl, apiConnected, apiError, liveTs, apiSourceInfo,
    // Datos procesados
    filteredRows, liveRecord, kpis, chartData,
    availableDays, availableMonths, availableYears,
    // Acciones
    loadFile, loadFromApi, testApiSource,
    toggleRefresh, stopRefresh,
    prevPeriod, nextPeriod, periodLabel,
  }
}