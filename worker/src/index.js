const JSON_HEADERS = {
  'Content-Type': 'application/json; charset=utf-8',
  'Cache-Control': 'no-store'
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/health' && request.method === 'GET') {
      return json({ ok: true, service: 'WPM RACER API proxy' });
    }

    if (url.pathname !== '/api') return json({ ok: false, error: 'Not found.' }, 404);

    const origin = request.headers.get('Origin');
    const allowedOrigins = String(env.ALLOWED_ORIGINS || '')
      .split(',')
      .map(function(value) { return value.trim().replace(/\/+$/, ''); })
      .filter(Boolean);
    if (!origin || allowedOrigins.indexOf(origin) < 0) {
      return json({ ok: false, error: 'Origin is not allowed.' }, 403, { Vary: 'Origin' });
    }

    const corsHeaders = {
      'Access-Control-Allow-Origin': origin,
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      'Vary': 'Origin'
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders });
    }
    if (request.method !== 'POST') {
      return json({ ok: false, error: 'Use POST for API requests.' }, 405, corsHeaders);
    }
    if (!/^application\/json(?:\s*;|$)/i.test(request.headers.get('Content-Type') || '')) {
      return json({ ok: false, error: 'Content-Type must be application/json.' }, 415, corsHeaders);
    }

    const contentLength = Number(request.headers.get('Content-Length') || 0);
    if (contentLength > 65536) {
      return json({ ok: false, error: 'Request body is too large.' }, 413, corsHeaders);
    }

    let body;
    let requestData;
    try {
      body = await request.text();
      if (new TextEncoder().encode(body).byteLength > 65536) {
        return json({ ok: false, error: 'Request body is too large.' }, 413, corsHeaders);
      }
      requestData = JSON.parse(body);
      if (!requestData || typeof requestData.action !== 'string' ||
          !requestData.payload || typeof requestData.payload !== 'object') {
        return json({ ok: false, error: 'Invalid API request.' }, 400, corsHeaders);
      }
    } catch (error) {
      return json({ ok: false, error: 'Request body must be valid JSON.' }, 400, corsHeaders);
    }

    if (!env.WPM_PROXY_SECRET) {
      return json({ ok: false, error: 'The backend proxy secret is not configured.' }, 503, corsHeaders);
    }

    let backendUrl;
    try {
      backendUrl = new URL(env.APPS_SCRIPT_URL);
      if (backendUrl.protocol !== 'https:' || backendUrl.hostname !== 'script.google.com' ||
          !/^\/macros\/s\/[^/]+\/exec$/.test(backendUrl.pathname)) {
        throw new Error('Invalid Apps Script deployment URL.');
      }
    } catch (error) {
      return json({ ok: false, error: 'The Apps Script backend URL is not configured correctly.' }, 503, corsHeaders);
    }

    try {
      const upstream = await fetch(backendUrl.toString(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ proxySecret: env.WPM_PROXY_SECRET, apiRequest: requestData }),
        redirect: 'follow'
      });
      const responseText = await upstream.text();
      let responseData;
      try {
        responseData = JSON.parse(responseText);
      } catch (error) {
        return json({ ok: false, error: 'Apps Script returned a non-JSON response.' }, 502, corsHeaders);
      }
      if (!upstream.ok) {
        return json({ ok: false, error: 'Apps Script backend request failed.' }, 502, corsHeaders);
      }
      return json(responseData, 200, corsHeaders);
    } catch (error) {
      return json({ ok: false, error: 'Could not reach the Apps Script backend.' }, 502, corsHeaders);
    }
  }
};

function json(data, status, extraHeaders) {
  const headers = new Headers(JSON_HEADERS);
  Object.keys(extraHeaders || {}).forEach(function(key) {
    headers.set(key, extraHeaders[key]);
  });
  return new Response(JSON.stringify(data), { status: status || 200, headers: headers });
}
