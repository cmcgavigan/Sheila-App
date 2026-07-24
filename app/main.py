"""FastAPI app — the whole server in one process.

Replaces v1's server.js + tray.js + watchdog.py + backup.py + launch scripts.
"""
from __future__ import annotations

import asyncio
import base64
import io

from fastapi import FastAPI, Request
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    JSONResponse,
    RedirectResponse,
    Response,
)
from fastapi.staticfiles import StaticFiles

from . import auth, config, db, images, jobs, netinfo
from .routes import router

PUBLIC_DIR = config.ROOT / "public"

app = FastAPI(title="Sheila PMU", version=config.APP_VERSION, docs_url=None, redoc_url=None)
app.include_router(router)

_public_url = ""  # resolved at startup


@app.middleware("http")
async def authentication_boundary(request: Request, call_next):
    try:
        if request.url.path.startswith("/api/") and int(
                request.headers.get("content-length", "0")) > 8 * 1024 * 1024:
            return JSONResponse({"error": "request body too large"}, status_code=413)
    except ValueError:
        return JSONResponse({"error": "invalid content length"}, status_code=400)
    if config.AUTH_DISABLED:
        return await call_next(request)
    path = request.url.path
    public = path in {"/login", "/api/auth/login", "/api/health", "/cert"} or path.startswith("/api/trip/automation/")
    if not public and not auth.session(request):
        if path.startswith("/api/"):
            return JSONResponse({"error": "Authentication required"}, status_code=401)
        if request.method == "GET":
            return RedirectResponse(f"/login?next={request.url.path}", status_code=303)
        return Response("Authentication required", status_code=401)
    if (path.startswith("/api/") and not path.startswith("/api/trip/automation/") and request.method not in {"GET", "HEAD", "OPTIONS"}
            and path not in {"/api/auth/login"} and not auth.csrf_ok(request, "")):
        return JSONResponse({"error": "CSRF token required"}, status_code=403)
    return await call_next(request)


@app.on_event("startup")
async def _startup():
    global _public_url
    db.init_db()
    with db.get_db() as conn:
        if not config.AUTH_DISABLED:
            auth.ensure_password(conn)
        images.recover_pending_photos(
            conn, config.PHOTOS_DIR, config.DATA_DIR / ".receipt-staging"
        )
    _public_url = netinfo.resolve_public_url()
    base = _public_url or f"https://localhost:{config.PORT}"
    try:
        (config.ROOT / "current-url.txt").write_text(base, encoding="utf-8")
    except OSError:
        pass
    asyncio.get_event_loop().create_task(jobs.background_loop())
    print(f"\n  Sheila PMU v{config.APP_VERSION} - running (HTTPS)")
    print("  -----------------------------------------")
    print(f"  Database:    {config.DB_PATH}")
    print(f"  Photos:      {config.PHOTOS_DIR}")
    print(f"  On laptop:   https://localhost:{config.PORT}")
    if _public_url:
        print(f"  Public URL:  {_public_url}    (valid cert, used by setup QR)")
    for ip in netinfo.get_local_ips():
        print(f"  fallback:    https://{ip}:{config.PORT}")
    print(f"\n  Phone setup (QR wall):  {base}/setup")
    print(f"  Treatments editor:      {base}/treatments\n")


@app.get("/login")
async def login_page(next: str = "/"):
    safe_next = next if next.startswith("/") and not next.startswith("//") else "/"
    return HTMLResponse(f"""<!doctype html><meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Sheila — Sign in</title>
<style>body{{font:17px system-ui;max-width:28rem;margin:15vh auto;padding:1rem;background:#0c0c0d;color:#f4efe6}}
input,button{{width:100%;padding:1rem;margin:.5rem 0;border-radius:.7rem;font-size:1rem}}
input{{background:#141412;color:#f4efe6;border:1px solid #3a342a}}button{{background:#c9a24b;border:0}}
#e{{color:#e0594f;min-height:1.5rem}}</style>
<h1>Sheila PMU</h1><p>Sign in to continue.</p>
<form id='f'><input id='p' type='password' autocomplete='current-password' autofocus>
<button>Sign in</button><div id='e'></div></form>
<script>f.onsubmit=async e=>{{e.preventDefault();let r=await fetch('/api/auth/login',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{password:p.value}})}});if(r.ok)location.href={safe_next!r};else document.querySelector('#e').textContent='Sign-in failed';}}</script>""")


@app.post("/api/auth/login")
async def auth_login(request: Request):
    body = await request.json()
    identity = request.client.host if request.client else "local"
    if auth.rate_limited(identity):
        return JSONResponse({"error": "Too many failed attempts"}, status_code=429)
    with db.get_db() as conn:
        if not config.AUTH_DISABLED:
            auth.ensure_password(conn)
        result = auth.login(conn, str(body.get("password") or ""), identity)
    if not result:
        return JSONResponse({"error": "Invalid credentials"}, status_code=401)
    token, csrf = result
    response = JSONResponse({"ok": True, "csrfToken": csrf})
    response.set_cookie(auth.SESSION_COOKIE, token, httponly=True, secure=config.AUTH_SECURE_COOKIE,
                        samesite="strict", max_age=8 * 3600, path="/")
    response.set_cookie(auth.CSRF_COOKIE, csrf, httponly=False, secure=config.AUTH_SECURE_COOKIE,
                        samesite="strict", max_age=8 * 3600, path="/")
    return response


@app.post("/api/auth/logout")
async def auth_logout(request: Request):
    token = request.cookies.get(auth.SESSION_COOKIE)
    auth._sessions.pop(token, None)
    response = JSONResponse({"ok": True})
    response.delete_cookie(auth.SESSION_COOKIE, path="/")
    response.delete_cookie(auth.CSRF_COOKIE, path="/")
    return response


# --------------------------------------------------------------------------
# Update endpoints (PWA "Update now" banner)
# --------------------------------------------------------------------------
@app.get("/api/update-check")
async def update_check():
    try:
        return jobs.update_status()
    except Exception as e:  # noqa: BLE001
        return {"mode": "error", "updateAvailable": False, "error": str(e)}


@app.post("/api/update")
async def do_update():
    result = jobs.perform_update()
    if "error" in result:
        return JSONResponse(result, status_code=500)
    if result.get("restarting"):
        jobs.schedule_restart()
    return result


# --------------------------------------------------------------------------
# Cert download (for manually trusting the self-signed cert if ever needed)
# --------------------------------------------------------------------------
@app.get("/cert")
async def cert():
    if not netinfo.CERT_PEM.exists():
        return Response("Certificate not generated yet.", status_code=404)
    return FileResponse(netinfo.CERT_PEM, filename="Sheila.crt",
                        media_type="application/x-x509-ca-cert")


# --------------------------------------------------------------------------
# /setup — the QR wall. Same three cards as v1: Tailscale, the app, treatments.
# --------------------------------------------------------------------------
def _qr_data_url(text: str) -> str:
    import qrcode

    img = qrcode.make(text, box_size=6, border=1)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


@app.get("/setup")
async def setup_page():
    base = (_public_url or f"https://localhost:{config.PORT}").rstrip("/")
    on_tailscale = ".ts.net" in base
    ts_play = "https://play.google.com/store/apps/details?id=com.tailscale.ipn"
    ts_qr, app_qr, treat_qr = (_qr_data_url(ts_play), _qr_data_url(base),
                               _qr_data_url(base + "/treatments"))

    def card(n, title, body):
        return f'<div class="card"><div class="num">{n}</div><h2>{title}</h2>{body}</div>'

    cards = card("1", "Install Tailscale on the phone",
        f"<p>Scan to get Tailscale, then <b>sign in with the SAME account as this laptop</b> "
        f"and toggle it <b>On</b>. This lets the phone reach the laptop on any network — "
        f"abroad, mobile data, hotel Wi-Fi.</p><img src='{ts_qr}' width='230' height='230'>"
        f"<p class='sub'>iPhone: search “Tailscale” in the App Store.</p>")
    cards += card("2", "Open Sheila’s Receipts app",
        f"<p>Scan to open the app{' (secure, no warnings)' if on_tailscale else ''}. Then "
        f"Safari/Chrome menu → <b>Add to Home Screen</b>, and tap <b>Allow</b> if it asks "
        f"about notifications.</p><img src='{app_qr}' width='230' height='230'>"
        f"<p class='sub'>In/Out/Trip toggle is at the top. No key, no typing.</p>")
    cards += card("3", "Treatments editor (optional)",
        f"<p>Scan to open the treatments list editor. Enter the PIN to add, edit or delete "
        f"treatments and their prices.</p><img src='{treat_qr}' width='230' height='230'>"
        f"<p class='sub'>Add this to the Home Screen too if you’ll edit prices often.</p>")

    warn = "" if on_tailscale else (
        '<div class="warn">Tailscale isn’t serving a certificate yet, so the phone may show '
        'a “not secure” warning. Finish Tailscale setup on the laptop (sign in + turn on '
        'HTTPS) and reload — the warning disappears.</div>')

    return HTMLResponse(f"""<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sheila — phone setup</title>
<style>
:root{{--ink:#0f1222;--muted:#6b7280;--line:#e5e7eb;--accent:#b6356b}}
*{{box-sizing:border-box}}
body{{font:16px/1.55 -apple-system,Segoe UI,Roboto,sans-serif;margin:0;color:var(--ink);background:#fdf2f8}}
.wrap{{max-width:60rem;margin:0 auto;padding:1.5rem 1.25rem 3rem}}
h1{{font-size:1.5rem;margin:.2rem 0}}.lead{{color:var(--muted);margin:.2rem 0 1.4rem}}
.grid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:1rem}}
@media(max-width:820px){{.grid{{grid-template-columns:1fr}}}}
.card{{position:relative;background:#fff;border:1px solid var(--line);border-radius:16px;padding:1.2rem 1.2rem 1.4rem}}
.card .num{{position:absolute;top:-12px;left:16px;width:30px;height:30px;border-radius:50%;background:var(--accent);color:#fff;font-weight:700;display:flex;align-items:center;justify-content:center}}
.card h2{{font-size:1.05rem;margin:.4rem 0 .5rem}}
.card p{{margin:.4rem 0}}.card .sub{{color:var(--muted);font-size:.85rem}}
.card img{{display:block;margin:.6rem auto 0;border-radius:12px;background:#fff;padding:6px;border:1px solid var(--line)}}
.warn{{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;border-radius:12px;padding:.8rem 1rem;margin:1rem 0}}
.done{{text-align:center;color:var(--muted);margin-top:1.6rem;font-size:.9rem}}
.addr{{font:13px ui-monospace,Consolas,monospace;color:var(--muted);word-break:break-all}}
</style>
<div class="wrap">
<h1>💅 Sheila — phone setup</h1>
<p class="lead">Scan these codes with the phone. The whole setup takes about a minute.</p>
{warn}
<div class="grid">{cards}</div>
<p class="done">App address: <span class="addr">{base}</span><br>After setup, always open from the Home-Screen icon.</p>
</div>""")


@app.get("/treatments")
async def treatments_page():
    return FileResponse(PUBLIC_DIR / "treatments.html")


# --------------------------------------------------------------------------
# /export — one-button page: generate the Steuerberater workbook, download it.
# --------------------------------------------------------------------------
@app.get("/export")
async def export_page():
    return HTMLResponse("""<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sheila — Excel-Export</title>
<style>
body{font:17px/1.5 -apple-system,Segoe UI,Roboto,sans-serif;background:#fdf2f8;color:#0f1222;
     display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}
.card{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:2rem;max-width:26rem;text-align:center}
button{background:#b6356b;color:#fff;border:0;border-radius:12px;padding:.9rem 1.6rem;font-size:1.05rem;font-weight:700;cursor:pointer}
button:disabled{opacity:.5}
.sub{color:#6b7280;font-size:.85rem;margin-top:1rem}
a{color:#b6356b;font-weight:700}
</style>
<div class="card">
<h1>📊 Steuerberater-Export</h1>
<p>Erstellt eine frische Excel-Datei aus allen Daten (Einnahmen, Ausgaben, Reisen) — §19, mit Abzugsfähigkeit.</p>
<button id="go">Export erstellen</button>
<p id="out" class="sub"></p>
<script>
document.getElementById('go').onclick = async () => {
  const b = document.getElementById('go'), o = document.getElementById('out');
  b.disabled = true; o.textContent = 'Wird erstellt…';
  try {
    const r = await fetch('/api/export', {method:'POST'});
    const j = await r.json();
    if (j.ok) { o.innerHTML = 'Fertig: <a href="' + j.download + '">' + j.file + '</a>'; }
    else { o.textContent = 'Fehler: ' + (j.error || r.status); }
  } catch (e) { o.textContent = 'Fehler: ' + e.message; }
  b.disabled = false;
};
</script>
</div>""")


# Static PWA — mounted last so /api, /setup etc. win.
app.mount("/", StaticFiles(directory=PUBLIC_DIR, html=True), name="public")
