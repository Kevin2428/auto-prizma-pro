from __future__ import annotations

import copy
import html
import io
import json
import os
import re
import threading
import uuid
import zipfile
from datetime import datetime, timezone

from fastapi import APIRouter, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response

import motor_prizma

from .automation import PrizmaTestAutomator
from .models import Severity, TestDocument, document_from_dict, document_to_dict
from .parser import parse_docx_bytes
from .serializer import serialize_test
from .validation import validate_document


router = APIRouter()
_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.RLock()
_AUTOMATION_JOBS: dict[str, dict] = {}
_AUTOMATION_LOCK = threading.RLock()
_AUTOMATION_SYNC: dict[str, dict] = {}
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"




def _ruta_token_prizma(usuario: str) -> str:
    tokens_dir = os.path.join(os.path.dirname(__file__), "..", "data", "tokens")
    os.makedirs(tokens_dir, exist_ok=True)
    return os.path.join(tokens_dir, f"prizma_{usuario}.json")


def _cargar_credenciales_prizma(usuario: str) -> dict:
    ruta = _ruta_token_prizma(usuario)
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"usuario": "", "recordar": True}


def _guardar_credenciales_prizma(usuario: str, creds: dict):
    ruta = _ruta_token_prizma(usuario)
    try:
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(creds, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _user(request: Request) -> str:
    return str(getattr(request.state, "usuario", "") or "").strip().lower()


def _job_for(request: Request, job_id: str) -> dict | None:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if not job or job.get("owner") != _user(request):
            return None
        return job


def _issue_counts(issues):
    return {
        "error": sum(issue.severity is Severity.ERROR for issue in issues),
        "warning": sum(issue.severity is Severity.WARNING for issue in issues),
    }


import os
import sys

def _es_google_conectado(usuario: str) -> bool:
    main_mod = sys.modules.get("main")
    if main_mod and hasattr(main_mod, "_google_conectado"):
        try:
            return bool(main_mod._google_conectado(usuario))
        except Exception:
            pass
    token_path = os.path.join(os.path.dirname(__file__), "..", "data", f"tokens_{usuario}.json")
    return os.path.exists(token_path)


def _patch_main_nav():
    """Garantiza en memoria que el botón Test aparezca en /cargue-actual sin tocar main.py."""
    main_mod = sys.modules.get("main")
    if not main_mod:
        return

    orig_vacio = getattr(main_mod, "_pagina_sin_cargue_actual", None)
    if orig_vacio and not getattr(orig_vacio, "_ap_patched", False):
        def nuevo_vacio(*args, **kwargs):
            resp = orig_vacio(*args, **kwargs)
            if hasattr(resp, "body"):
                body_str = resp.body.decode("utf-8")
                if 'href="/tests"' not in body_str and '<a class="nav-item activo" href="/cargue-actual">' in body_str:
                    body_str = body_str.replace(
                        '<a class="nav-item activo" href="/cargue-actual">',
                        '<a class="nav-item" href="/tests">▣ <span>Test</span></a><a class="nav-item activo" href="/cargue-actual">'
                    )
                    return HTMLResponse(body_str, status_code=resp.status_code)
            return resp
        nuevo_vacio._ap_patched = True
        setattr(main_mod, "_pagina_sin_cargue_actual", nuevo_vacio)

    orig_cargue = getattr(main_mod, "cargue_actual", None)
    if orig_cargue and not getattr(orig_cargue, "_ap_patched", False):
        def nuevo_cargue(*args, **kwargs):
            resp = orig_cargue(*args, **kwargs)
            if hasattr(resp, "body"):
                body_str = resp.body.decode("utf-8")
                if 'href="/tests"' not in body_str and '<a class="nav-item activo" href="/cargue-actual">' in body_str:
                    body_str = body_str.replace(
                        '<a class="nav-item activo" href="/cargue-actual">',
                        '<a class="nav-item" href="/tests">▣ <span>Test</span></a><a class="nav-item activo" href="/cargue-actual">'
                    )
                    return HTMLResponse(body_str, status_code=resp.status_code)
            return resp
        nuevo_cargue._ap_patched = True
        app = getattr(main_mod, "app", None)
        if app:
            for route in app.routes:
                if getattr(route, "path", None) == "/cargue-actual":
                    route.endpoint = nuevo_cargue
        setattr(main_mod, "cargue_actual", nuevo_cargue)


_patch_main_nav()


def _layout(title: str, body: str, active: str = "Test") -> str:
    _patch_main_nav()
    nav = "".join(
        f'<a class="nav-item {"activo" if label == active else ""}" href="{href}">{icon}<span>{label}</span></a>'
        for label, href, icon in [
            ("Inicio", "/", "⌂"),
            ("Test", "/tests", "▣"),
            ("Cargue actual", "/cargue-actual", "⇧"),
            ("Historial", "/historial", "◷"),
            ("Reportes", "/reportes", "▥"),
            ("Mi contraseña", "/cambiar-clave", "✎"),
            ("Salir", "/salir", "⏻"),
        ]
    )
    return f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)} - Auto Prizma Pro</title><style>
:root{{--fondo:#f7f8fc;--texto:#101828;--muted:#667085;--borde:#e5e7ef;--morado:#5548e8;--verde:#087f5b;--rojo:#b42318;--amarillo:#9a6700;--sombra:0 10px 30px rgba(29,41,57,.05)}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--fondo);color:var(--texto);font-family:Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}}
.app{{min-height:100vh;display:grid;grid-template-columns:235px 1fr}}
.sidebar{{position:sticky;top:0;height:100vh;background:#fff;border-right:1px solid var(--borde);padding:28px 20px;display:flex;flex-direction:column}}
.marca{{display:flex;align-items:center;gap:12px;margin-bottom:34px}}
.marca strong{{display:block;font-size:18px}}
.marca span{{display:block;color:var(--muted);font-size:12px;margin-top:3px}}
.logo{{width:44px;height:44px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(145deg,#6d5dfc,#4338ca);color:#fff;font-weight:900;font-size:22px;box-shadow:0 8px 20px rgba(79,70,229,.25)}}
.nav{{display:grid;gap:8px}}
.nav-item{{padding:12px 14px;border-radius:11px;color:#475467;font-size:14px;display:flex;gap:11px;align-items:center;text-decoration:none}}
.nav-item.activo,.nav-item:hover{{background:#f1efff;color:#4f46e5;font-weight:700}}
.contenido{{padding:34px 38px;min-width:0}}
.layout-tests{{display:grid;grid-template-columns:1fr 340px;gap:24px;align-items:start}}
.panel{{background:#fff;border:1px solid var(--borde);border-radius:18px;box-shadow:var(--sombra);padding:26px;margin-bottom:18px}}
.panel-ayuda{{padding:24px;background:#fff;border:1px solid var(--borde);border-radius:18px}}
.panel-ayuda h3{{margin:0 0 18px;font-size:17px;color:#101828}}
.paso-ayuda{{display:grid;grid-template-columns:30px 1fr;gap:12px;align-items:start;margin-bottom:16px;color:#475467;font-size:13px;line-height:1.5}}
.paso-ayuda b{{width:28px;height:28px;border-radius:50%;background:#f1efff;color:#4f46e5;display:grid;place-items:center;font-size:13px;font-weight:800}}
.ayuda-separador{{height:1px;background:var(--borde);margin:22px 0}}
.titulo-consejos{{color:#4f46e5;font-size:16px;margin:0 0 12px}}
.consejo{{color:#475467;font-size:13px;line-height:1.55;margin:0 0 10px}}
h1{{margin:0 0 8px;font-size:29px}}
h2{{margin:0 0 12px;font-size:20px}}
p{{color:var(--muted);line-height:1.5;margin:0 0 14px}}
.pasos{{display:flex;gap:8px;flex-wrap:wrap;margin:18px 0 24px}}
.paso{{padding:8px 14px;border-radius:999px;background:#f1efff;color:#5548e8;font-size:12px;font-weight:800}}
.dropzone{{display:grid;place-items:center;border:2px dashed #b8b3f8;border-radius:14px;padding:32px 20px;text-align:center;background:#fbfaff;cursor:pointer;transition:all .2s}}
.dropzone:hover,.dropzone.dragover{{border-color:#5548e8;background:#f4f2ff}}
input,textarea,select{{width:100%;border:1px solid #d0d5dd;border-radius:10px;padding:11px 13px;font:inherit;color:inherit;background:#fff}}
textarea{{min-height:78px;resize:vertical}}
label{{display:block;font-size:12px;font-weight:800;margin:13px 0 6px;color:#344054}}
button,.boton{{display:inline-flex;align-items:center;justify-content:center;gap:8px;border:0;border-radius:10px;padding:12px 18px;background:var(--morado);color:#fff;font-weight:800;font-size:14px;text-decoration:none;cursor:pointer}}
button:hover,.boton:hover{{opacity:.93}}
button.secundario,.boton.secundario{{background:#fff;color:#5548e8;border:1px solid #c7c2ff}}
.acciones{{display:flex;gap:12px;flex-wrap:wrap;margin-top:22px}}
.test-card{{border:1px solid var(--borde);border-radius:16px;padding:22px;margin-top:18px;background:#fcfcff;box-shadow:0 4px 12px rgba(0,0,0,.02)}}
.section{{border-left:4px solid #7c6dfc;padding-left:16px;margin:20px 0}}
.question{{background:#fff;border:1px solid var(--borde);border-radius:12px;padding:16px;margin:14px 0;transition:all .2s;scroll-margin-top:24px}}
.question.question-error{{border:2px solid #f04438 !important;background:#fffbfa !important;box-shadow:0 0 0 4px rgba(240,68,56,0.12);}}
.badge{{display:inline-block;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:800;margin:3px 6px 3px 0}}
.badge.ok{{background:#ecfdf3;color:var(--verde)}}
.badge.warn{{background:#fff8e1;color:var(--amarillo)}}
.badge.error{{background:#fef3f2;color:var(--rojo)}}
.issue{{padding:12px 14px;border-radius:10px;margin:8px 0;font-size:13px;line-height:1.5}}
.issue.error{{background:#fef3f2;color:var(--rojo);border:1px solid #fecdca}}
.issue.warning{{background:#fff8e1;color:var(--amarillo);border:1px solid #fedf89}}
.issue.ok{{background:#ecfdf3;color:var(--verde);border:1px solid #a6f4c5}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.option-row,.pair-row{{display:grid;grid-template-columns:42px 1fr auto;gap:10px;align-items:center;margin-top:8px}}
.pair-row{{grid-template-columns:1fr 1fr}}
.muted{{color:var(--muted);font-size:12px}}
@media(max-width:1100px){{.layout-tests{{grid-template-columns:1fr}}.panel-ayuda{{order:2}}}}
@media(max-width:900px){{.app{{grid-template-columns:1fr}}.sidebar{{display:none}}.contenido{{padding:20px}}.grid2{{grid-template-columns:1fr}}}}
.estado-servicio{{margin-top:auto;border:1px solid var(--borde);border-radius:14px;padding:15px}}
.servicio-linea{{font-size:12px;font-weight:800;color:#07894f;margin-bottom:14px}}
.punto{{width:8px;height:8px;background:#12b76a;border-radius:50%;display:inline-block;margin-right:7px}}
.servicio-mini{{display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--muted);margin-top:10px}}
.chip{{background:#eef2ff;color:#4f46e5;border-radius:999px;padding:4px 8px}}
html[data-theme="dark"] .test-card,
html[data-theme="dark"] .panel-ayuda,
html[data-theme="dark"] .question,
html[data-theme="dark"] .dropzone {{
  background: #111827 !important;
  color: #e5e7eb !important;
  border-color: #263244 !important;
}}
html[data-theme="dark"] .question.question-error {{
  background: #2a1a1a !important;
  border-color: #f04438 !important;
}}
html[data-theme="dark"] .section {{ border-left-color: #6d5dfc !important; }}
html[data-theme="dark"] .titulo-consejos,
html[data-theme="dark"] .paso-ayuda,
html[data-theme="dark"] .consejo,
html[data-theme="dark"] .muted {{ color: #aeb9ca !important; }}
html[data-theme="dark"] .ayuda-separador {{ background: #263244 !important; }}
html[data-theme="dark"] .issue.error {{ background: #2a1a1a !important; border-color: #633 !important; }}
html[data-theme="dark"] .issue.warning {{ background: #2a2210 !important; border-color: #7a5c00 !important; }}
html[data-theme="dark"] .issue.ok {{ background: #103322 !important; border-color: #0a5c3a !important; }}
html[data-theme="dark"] .badge.ok {{ background: #103322 !important; }}
html[data-theme="dark"] .badge.warn {{ background: #2a2210 !important; }}
html[data-theme="dark"] .badge.error {{ background: #2a1a1a !important; }}

/* ---- Estados dinamicos reutilizables (banners, kpis, badges, tarjeta en vivo) ---- */
.status-banner{{display:flex;gap:14px;align-items:flex-start;margin-bottom:20px;border-radius:14px;padding:20px;border:2px solid transparent;box-shadow:0 4px 12px rgba(0,0,0,.06)}}
.status-banner .status-icon{{width:40px;height:40px;border-radius:50%;display:grid;place-items:center;font-size:22px;font-weight:900;flex-shrink:0}}
.status-banner h3{{margin:0 0 6px;font-size:17px;font-weight:800}}
.status-banner p{{margin:0 0 12px;font-size:14px;line-height:1.5}}
.status-banner .acciones{{margin-top:0}}
.status-banner--ok{{background:#f6fef9;border-color:#a6f4c5}}
.status-banner--ok .status-icon{{background:#d1fadf;color:#027a48}}
.status-banner--ok h3,.status-banner--ok p{{color:#027a48}}
.status-banner--error{{background:#fffbfa;border-color:#fecdca}}
.status-banner--error .status-icon{{background:#fee4e2;color:#d92d20}}
.status-banner--error h3{{color:#b42318}}
.status-banner--error p{{color:#7a271a}}
.status-banner--warn{{background:#fffcf5;border-color:#fedf89}}
.status-banner--warn .status-icon{{background:#fef0c7;color:#b54708}}
.status-banner--warn h3{{color:#92400e}}
.status-banner--warn p{{color:#78350f}}

.kpi-card{{background:#fff;border:1px solid var(--borde);border-radius:12px;padding:16px;box-shadow:0 1px 3px rgba(16,24,40,.05)}}
.kpi-card .kpi-label{{font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted)}}
.kpi-card .kpi-valor{{font-size:28px;font-weight:900;margin-top:4px;color:var(--texto)}}
.kpi-card--ok{{background:#f6fef9;border-color:#a6f4c5}}
.kpi-card--ok .kpi-label,.kpi-card--ok .kpi-valor{{color:#027a48}}
.kpi-card--warn{{background:#fffcf5;border-color:#fedf89}}
.kpi-card--warn .kpi-label,.kpi-card--warn .kpi-valor{{color:#b54708}}
.kpi-card--error{{background:#fffbfa;border-color:#fecdca}}
.kpi-card--error .kpi-label,.kpi-card--error .kpi-valor{{color:#b42318}}

.badge-dinamico{{display:inline-block;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:800}}
.badge-dinamico--pendiente{{background:#f2f4f7;color:#344054}}
.badge-dinamico--proceso{{background:#fef0c7;color:#b54708}}
.badge-dinamico--ok{{background:#ecfdf3;color:#087f5b}}
.badge-dinamico--error{{background:#fef3f2;color:#b42318}}
.badge-dinamico--detenido{{background:#e5e7eb;color:#374151}}

.live-card{{border-radius:16px;padding:20px 24px;box-shadow:0 4px 14px rgba(0,0,0,.04);border:2px solid transparent;margin-bottom:24px;transition:all .3s ease}}
.live-card--running{{background:#fbfaff;border-color:#6366f1}}
.live-card--waiting{{background:#fef2f2;border-color:#ef4444}}
.live-card h3{{margin:0;font-size:16px;font-weight:800;color:var(--texto)}}
.live-card p{{margin:3px 0 0;font-size:13px;color:var(--muted)}}

.progress-track{{background:#e2e8f0;border-radius:999px;height:9px;overflow:hidden}}
.progress-fill{{background:linear-gradient(90deg,#5548e8,#818cf8);height:100%;transition:width .4s ease}}

.mode-option{{flex:1;min-width:180px;border:1px solid var(--borde);border-radius:12px;padding:14px;cursor:pointer;background:#fff;transition:all .2s}}
.mode-option strong{{margin-left:7px;font-size:14px}}
.mode-option span{{display:block;margin:5px 0 0 25px;color:var(--muted);font-size:12px}}
.mode-option--activo{{border-color:#5548e8;background:#f8f7ff}}

.alert-soft{{margin-top:14px;border-radius:12px;padding:14px;display:flex;gap:12px;align-items:center}}
.alert-soft--ok{{background:#ecfdf3;border:1px solid #a6f4c5;color:#027a48}}
.alert-soft--warn{{background:#fff8df;border:1px solid #f5c451;color:#7a4b00;align-items:flex-start}}
.alert-soft a{{color:#4f46e5;font-weight:700}}

.test-card-header{{background:#f8f9fc;border-bottom:1px solid #edeafc}}
.chip-morado{{display:inline-block;font-size:12px;background:#edeafc;padding:3px 8px;border-radius:6px;color:#4338ca;font-weight:600}}

#btn-guardar-correcciones.is-dirty{{background:#e11d48 !important;color:#fff !important;font-weight:800;box-shadow:0 0 0 4px rgba(225,29,72,.35)}}

html[data-theme="dark"] .status-banner--ok{{background:#0b2318 !important;border-color:#0a5c3a !important}}
html[data-theme="dark"] .status-banner--ok h3,html[data-theme="dark"] .status-banner--ok p{{color:#86efac !important}}
html[data-theme="dark"] .status-banner--error{{background:#2a1a1a !important;border-color:#633 !important}}
html[data-theme="dark"] .status-banner--error h3{{color:#fca5a5 !important}}
html[data-theme="dark"] .status-banner--error p{{color:#fecaca !important}}
html[data-theme="dark"] .status-banner--warn{{background:#2a2210 !important;border-color:#7a5c00 !important}}
html[data-theme="dark"] .status-banner--warn h3{{color:#fbbf24 !important}}
html[data-theme="dark"] .status-banner--warn p{{color:#fde68a !important}}

html[data-theme="dark"] .kpi-card{{background:#111827 !important;border-color:#263244 !important}}
html[data-theme="dark"] .kpi-card .kpi-label{{color:#aeb9ca !important}}
html[data-theme="dark"] .kpi-card .kpi-valor{{color:#f3f4f6 !important}}
html[data-theme="dark"] .kpi-card--ok{{background:#0b2318 !important;border-color:#0a5c3a !important}}
html[data-theme="dark"] .kpi-card--ok .kpi-label,html[data-theme="dark"] .kpi-card--ok .kpi-valor{{color:#4ade80 !important}}
html[data-theme="dark"] .kpi-card--warn{{background:#2a2210 !important;border-color:#7a5c00 !important}}
html[data-theme="dark"] .kpi-card--warn .kpi-label,html[data-theme="dark"] .kpi-card--warn .kpi-valor{{color:#fbbf24 !important}}
html[data-theme="dark"] .kpi-card--error{{background:#2a1a1a !important;border-color:#633 !important}}
html[data-theme="dark"] .kpi-card--error .kpi-label,html[data-theme="dark"] .kpi-card--error .kpi-valor{{color:#fca5a5 !important}}

html[data-theme="dark"] .badge-dinamico--pendiente{{background:#1f2937 !important;color:#cbd5e1 !important}}
html[data-theme="dark"] .badge-dinamico--proceso{{background:#2a2210 !important;color:#fbbf24 !important}}
html[data-theme="dark"] .badge-dinamico--ok{{background:#0b2318 !important;color:#4ade80 !important}}
html[data-theme="dark"] .badge-dinamico--error{{background:#2a1a1a !important;color:#fca5a5 !important}}
html[data-theme="dark"] .badge-dinamico--detenido{{background:#263244 !important;color:#cbd5e1 !important}}

html[data-theme="dark"] .live-card--running{{background:#111827 !important;border-color:#6d5dfc !important}}
html[data-theme="dark"] .live-card--waiting{{background:#2a1a1a !important;border-color:#ef4444 !important}}

html[data-theme="dark"] .mode-option{{background:#111827 !important;border-color:#263244 !important;color:#e5e7eb !important}}
html[data-theme="dark"] .mode-option span{{color:#aeb9ca !important}}
html[data-theme="dark"] .mode-option--activo{{background:#1d2540 !important;border-color:#818cf8 !important}}

html[data-theme="dark"] .alert-soft--ok{{background:#0b2318 !important;border-color:#0a5c3a !important;color:#86efac !important}}
html[data-theme="dark"] .alert-soft--warn{{background:#2a2210 !important;border-color:#7a5c00 !important;color:#fde68a !important}}
html[data-theme="dark"] .alert-soft a{{color:#a5b4fc !important}}

html[data-theme="dark"] .progress-track{{background:#1f2937 !important}}

html[data-theme="dark"] .test-card-header{{background:#0f172a !important;border-bottom-color:#263244 !important}}
html[data-theme="dark"] .chip-morado{{background:#1d2540 !important;color:#a5b4fc !important}}
html[data-theme="dark"] button.secundario,
html[data-theme="dark"] .boton.secundario{{background:#111827 !important;color:#a5b4fc !important;border-color:#334155 !important}}
.boton.secundario.peligro{{border-color:#fda29b;color:#b42318}}
html[data-theme="dark"] .boton.secundario.peligro{{background:#2a1a1a !important;color:#fca5a5 !important;border-color:#7a3b3b !important}}
</style><link rel="stylesheet" href="/estilos-responsive.css"><script src="/tema.js" defer></script></head><body><div class="app"><aside class="sidebar"><div class="marca"><div class="logo">A</div><div><strong>Auto Prizma Pro</strong><span>Automatización PRIZMA</span></div></div><nav class="nav">{nav}</nav><div class="estado-servicio"><div class="servicio-linea"><span class="punto"></span> Servicio activo</div><div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div><div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div></div></aside><main class="contenido">{body}</main></div></body></html>'''


def _render_issues(issues) -> str:
    if not issues:
        return '<span class="badge ok" id="badge-todas-validadas">✓ Todas las preguntas y claves validadas</span>'
    items = []
    for issue in issues:
        m = re.match(r"tests\[(\d+)\]\.sections\[(\d+)\]\.questions\[(\d+)\]", issue.path)
        if m:
            t_idx, s_idx, q_idx = m.groups()
            anchor = f"#pregunta-{t_idx}-{s_idx}-{q_idx}"
            target_id = f"pregunta-{t_idx}-{s_idx}-{q_idx}"
            items.append(
                f'<div class="issue {issue.severity.value}" data-issue-target="{target_id}" data-original-msg="{html.escape(issue.message, quote=True)}" style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">'
                f'<div class="issue-text"><strong>{issue.severity.value.title()}</strong>: {html.escape(issue.message)}</div>'
                f'<a class="issue-link" href="{anchor}" onclick="irAPregunta(\'{target_id}\', {t_idx}); return false;" style="color:#5548e8;font-weight:700;text-decoration:underline;font-size:12px;cursor:pointer;">Ir a la pregunta ➔</a>'
                f'</div>'
            )
        else:
            items.append(f'<div class="issue {issue.severity.value}"><div class="issue-text"><strong>{issue.severity.value.title()}</strong>: {html.escape(issue.message)}</div></div>')
    return "".join(items)


def _render_review(job_id: str, job: dict, message: str = "") -> str:
    document: TestDocument = job["document"]
    issues = validate_document(document)
    counts = _issue_counts(issues)
    state = json.dumps(document_to_dict(document), ensure_ascii=False).replace("<", "\\u003c")
    issues_by_path: dict[str, list] = {}
    for issue in issues:
        issues_by_path.setdefault(issue.path, []).append(issue)

    owner = job.get("owner", "")
    creds = _cargar_credenciales_prizma(owner)

    job_creds = job.get("credenciales_prizma") or {}
    usuario_prizma_val = job_creds.get("usuario") or creds.get("usuario", "")
    clave_prizma_val = job_creds.get("clave", "")

    cascada_datos = job.get("cascada_datos") or {}
    cascada_facultad_val = cascada_datos.get("facultad", "")
    cascada_programa_val = cascada_datos.get("programa") or document.program or ""
    cascada_pensum_val = cascada_datos.get("pensum", "")
    cascada_nivel_val = cascada_datos.get("nivel", "")
    cascada_asignatura_val = cascada_datos.get("asignatura", "")


    programa_sugerido = cascada_programa_val or document.program or ""
    primera_semana = document.tests[0].week if document.tests else ""
    primer_corte = document.tests[0].cut if document.tests else ""

    cards = []
    for test_index, test in enumerate(document.tests):
        sections = []
        for section_index, section in enumerate(test.sections):
            questions = []
            for question_index, question in enumerate(section.questions):
                q_path = f"tests[{test_index}].sections[{section_index}].questions[{question_index}]"
                q_issues = issues_by_path.get(q_path, [])
                has_q_error = any(i.severity == Severity.ERROR for i in q_issues)
                error_class = " question-error" if has_q_error else ""
                
                error_alert = ""
                if q_issues:
                    error_items = "".join(f'<li style="margin:3px 0;">{html.escape(i.message)}</li>' for i in q_issues)
                    error_alert = f'<div class="issue error" style="margin:0 0 12px 0;"><strong>⚠️ Corrección requerida en Pregunta {question_index + 1}:</strong><ul style="margin:4px 0 0 18px;padding:0;">{error_items}</ul></div>'
                
                badge_estado = '<span class="badge error">❌ Requiere atención</span>' if has_q_error else '<span class="badge ok">✓ Válida</span>'

                option_html = ""
                if section.code == "trueOrFalse":
                    value = "true" if question.answer is True else "false" if question.answer is False else ""
                    option_html = f'<div style="margin-top:10px;"><label>Respuesta Correcta</label><select style="max-width:200px;" data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-field="answer"><option value="">Selecciona una opción</option><option value="true" {"selected" if value == "true" else ""}>Verdadero</option><option value="false" {"selected" if value == "false" else ""}>Falso</option></select></div>'
                elif section.code == "multipleChoice":
                    option_html = "<div style='margin-top:10px;'><label>Opciones y Clave (marca la casilla de la respuesta correcta)</label>" + "".join(
                        f'<div class="option-row"><strong>{html.escape(option.letter.upper())}</strong><input data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-option="{option_index}" data-field="option_text" value="{html.escape(option.text, quote=True)}"><label style="margin:0;display:flex;align-items:center;gap:5px;cursor:pointer;"><input type="checkbox" style="width:auto;" data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-option="{option_index}" data-field="option_correct" {"checked" if option.letter in question.correct_option_letters else ""}> Correcta</label></div>'
                        for option_index, option in enumerate(question.options)
                    ) + "</div>"
                elif section.code == "matching":
                    option_html = "<div style='margin-top:10px;'><label>Parejas Concepto - Definición</label>" + "".join(
                        f'<div class="pair-row"><input placeholder="Concepto (Col A)" data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-pair="{pair_index}" data-field="pair_left" value="{html.escape(pair.left, quote=True)}"><input placeholder="Descripción (Col B)" data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-pair="{pair_index}" data-field="pair_right" value="{html.escape(pair.right, quote=True)}"></div>'
                        for pair_index, pair in enumerate(question.pairs)
                    ) + "</div>"
                questions.append(
                    f'<div class="question{error_class}" id="pregunta-{test_index}-{section_index}-{question_index}">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">'
                    f'<label style="margin:0;font-weight:800;font-size:13px;">Pregunta {question_index + 1}</label>'
                    f'{badge_estado}'
                    f'</div>'
                    f'{error_alert}'
                    f'<textarea data-test="{test_index}" data-section="{section_index}" data-question="{question_index}" data-field="prompt">{html.escape(question.prompt)}</textarea>'
                    f'{option_html}'
                    f'</div>'
                )
            sections.append(f'<div class="section"><h3>{html.escape(section.name)} ({len(section.questions)} pregunta{"s" if len(section.questions)!=1 else ""})</h3>{"".join(questions)}</div>')
        num_preguntas_test = sum(len(sec.questions) for sec in test.sections)
        cards.append(f'''<article class="test-card" id="test-card-{test_index}" style="margin-bottom:16px;padding:0;overflow:hidden;">
            <div class="test-card-header" onclick="toggleTestCard({test_index})" style="display:flex;justify-content:space-between;align-items:center;padding:14px 18px;cursor:pointer;user-select:none;">
                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                    <span class="badge ok" style="font-weight:700;">Test {test_index + 1} de {len(document.tests)}</span>
                    <strong style="font-size:15px;">{html.escape(test.title or f"Evaluación {test_index + 1}")}</strong>
                    <span class="chip-morado">{num_preguntas_test} pregunta{'s' if num_preguntas_test != 1 else ''}</span>
                    <span class="muted" style="font-size:12px;">Semana {html.escape(test.week or "N/A")} | Corte {html.escape(test.cut or "N/A")}</span>
                </div>
                <div style="display:flex;align-items:center;gap:8px;">
                    <span id="test-toggle-text-{test_index}" style="font-size:12px;font-weight:700;color:#6366f1;">{"Contraer ▲" if test_index == 0 else "Desplegar ▼"}</span>
                </div>
            </div>
            <div class="test-card-body" id="test-card-body-{test_index}" style="padding:18px;{'display:block;' if test_index == 0 else 'display:none;'}">
                <div class="grid2">
                    <div><label>Título del test</label><input data-test="{test_index}" data-field="title" value="{html.escape(test.title, quote=True)}"></div>
                    <div class="grid2">
                        <div><label>Semana</label><input data-test="{test_index}" data-field="week" value="{html.escape(test.week, quote=True)}"></div>
                        <div><label>Corte</label><input data-test="{test_index}" data-field="cut" value="{html.escape(test.cut, quote=True)}"></div>
                    </div>
                </div>
                <label>Indicaciones / Descripción</label>
                <textarea data-test="{test_index}" data-field="description">{html.escape(test.description)}</textarea>
                {"".join(sections)}
            </div>
        </article>''')
    message_html = f'<div class="issue ok">{html.escape(message)}</div>' if message else ''
    error_badge_style = "" if counts["error"] else "display:none;"
    ok_badge_style = "" if not counts["error"] else "display:none;"

    body = f'''<h1>Revisión de Tests</h1><p>Analiza, audita y corrige preguntas antes de conectar la plataforma PRIZMA.</p>
    <div class="pasos"><span class="paso">1 Entrada ✓</span><span class="paso">2 Revisión activa</span><span class="paso">3 Configuración PRIZMA</span><span class="paso">4 Carga en cola</span></div>
    {message_html}
    <section class="panel">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:12px;">
            <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                <span class="badge ok">{len(document.tests)} test(s) detectado(s)</span>
                <span class="badge error" id="badge-contador-errores" style="{error_badge_style}"><span id="num-errores">{counts['error']}</span> error(es) pendiente(s)</span>
                <span class="badge ok" id="badge-sin-errores" style="{ok_badge_style}">✓ Todas las preguntas y claves validadas</span>
                <span class="badge {'warn' if counts['warning'] else 'ok'}">{counts['warning']} advertencia(s)</span>
            </div>
            <a href="#bloque-acceso-prizma" class="boton" style="text-decoration:none;padding:8px 16px;font-size:13px;">
                ⚡ Ir a Carga en PRIZMA ↓
            </a>
        </div>
        <div id="contenedor-issues" style="margin-top:14px">{_render_issues(issues)}</div>
    </section>
    <form id="review-form" class="panel" method="post" action="/tests/{job_id}/guardar">
        <input type="hidden" name="payload" id="payload">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:10px;">
            <div style="display:flex;gap:8px;">
                <button type="button" class="boton secundario" onclick="expandirTodosTests()" style="font-size:12px;padding:6px 12px;cursor:pointer;">＋ Expandir todos</button>
                <button type="button" class="boton secundario" onclick="contraerTodosTests()" style="font-size:12px;padding:6px 12px;cursor:pointer;">－ Contraer todos</button>
            </div>
        </div>
        {"".join(cards)}
        <div class="acciones">
            <button type="submit" id="btn-guardar-correcciones" style="transition:all .3s ease;padding:12px 22px;border-radius:10px;font-weight:700;">💾 Guardar correcciones</button>
            <a class="boton secundario" href="/tests/{job_id}/descargar">⬇ Descargar respaldo (JSON)</a>
            <a class="boton secundario" href="/tests">＋ Nuevo análisis</a>
        </div>
    </form>

    <section id="bloque-acceso-prizma" class="panel" style="margin-top:28px;border:2px solid var(--morado);background:#faf9ff;scroll-margin-top:24px;">
        <div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">
            <div style="width:42px;height:42px;border-radius:12px;background:var(--morado);color:#fff;display:grid;place-items:center;font-size:22px;flex-shrink:0;">🛡</div>
            <div>
                <h2 style="margin:0;font-size:20px;">Fase 3: Publicación y Carga en PRIZMA</h2>
                <p style="margin:3px 0 0;font-size:13px;">Automatización con Playwright en tiempo real contra <strong>admin.prizma.site</strong></p>
            </div>
        </div>

        <form id="form-prizma" method="post" action="/tests/{job_id}/ejecutar-prizma">
            <div class="grid2">
                <div>
                    <label>Usuario PRIZMA</label>
                    <input name="prizma_usuario" placeholder="Ej. 1020304050" value="{html.escape(usuario_prizma_val)}" required>
                </div>
                <div>
                    <label>Contraseña PRIZMA</label>
                    <input type="password" name="prizma_clave" placeholder="••••••••" value="{html.escape(clave_prizma_val)}" required>
                </div>
            </div>

            <div style="display:flex;gap:18px;margin:14px 0 18px;flex-wrap:wrap;align-items:center;">
                <label style="display:inline-flex;align-items:center;gap:8px;cursor:pointer;font-weight:600;font-size:13px;">
                    <input type="checkbox" name="recordar_credenciales" value="1" style="width:auto;" {"checked" if creds.get("recordar", True) else ""}>
                    Recordar usuario en este equipo
                </label>
            </div>

            <div style="border-top:1px dashed var(--borde);padding-top:16px;margin-top:10px;">
                <h3 style="margin:0 0 12px;font-size:15px;">Parámetros del Curso en PRIZMA (Selectores en Cascada)</h3>
                <div class="grid2">
                    <div>
                        <label>Facultad <small style="font-weight:normal;">(Opcional)</small></label>
                        <input name="cascada_facultad" placeholder="Opcional - Facultad institucional (ej. Virtual - Derecho)" value="{html.escape(cascada_facultad_val)}">
                    </div>
                    <div>
                        <label>Programa</label>
                        <input name="cascada_programa" placeholder="Código o Nombre (ej. COD111)" value="{html.escape(cascada_programa_val)}">
                    </div>
                </div>
                <div class="grid2">
                    <div>
                        <label>Pensum</label>
                        <input name="cascada_pensum" placeholder="Código del pensum (ej. COD111 o Pensum 2024)" value="{html.escape(cascada_pensum_val)}">
                    </div>
                    <div>
                        <label>Nivel de pensum</label>
                        <input name="cascada_nivel" placeholder="Nivel (ej. 3, Tres, Semestre 3...)" value="{html.escape(cascada_nivel_val)}">
                    </div>
                </div>
                <div class="grid2">
                    <div>
                        <label>Asignatura por pensum</label>
                        <input name="cascada_asignatura" placeholder="Código o Nombre de la asignatura" value="{html.escape(cascada_asignatura_val)}">
                    </div>
                    <div class="grid2">
                        <div>
                            <label>Semana</label>
                            <input name="cascada_semana" placeholder="Semana" value="{html.escape(primera_semana)}">
                        </div>
                        <div>
                            <label>Corte</label>
                            <input name="cascada_corte" placeholder="Corte" value="{html.escape(primer_corte)}">
                        </div>
                    </div>
                </div>
            </div>

            <div class="acciones" style="margin-top:22px;">
                <button type="submit" id="btn-lanzar-prizma" style="background:#087f5b;color:#fff;padding:14px 28px;font-size:15px;font-weight:800;border:none;border-radius:10px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;box-shadow:0 4px 12px rgba(8,127,91,0.25);transition:all .2s;">
                    ⚡ Crear y Publicar Evaluación en PRIZMA
                </button>
            </div>
        </form>
    </section>

    <script>
    const initial={state};
    function val(q){{return q.type==='checkbox'?q.checked:q.value}}

    function verificarPreguntaValida(qBox) {{
        const selAnswer = qBox.querySelector('[data-field="answer"]');
        if (selAnswer) {{
            return selAnswer.value === 'true' || selAnswer.value === 'false';
        }}
        const checkboxes = qBox.querySelectorAll('[data-field="option_correct"]');
        if (checkboxes.length > 0) {{
            return Array.from(checkboxes).some(cb => cb.checked);
        }}
        const promptArea = qBox.querySelector('[data-field="prompt"]');
        if (promptArea) {{
            return promptArea.value.trim().length > 0;
        }}
        return true;
    }}

    function actualizarEstadoPreguntaEnVivo(qBox) {{
        if (!qBox) return;
        const targetId = qBox.id;
        const issueTop = document.querySelector(`[data-issue-target="${{targetId}}"]`);
        const esValida = verificarPreguntaValida(qBox);

        if (esValida) {{
            qBox.classList.remove('question-error');
            const errAlert = qBox.querySelector('.issue.error');
            if (errAlert) errAlert.style.display = 'none';
            const badge = qBox.querySelector('.badge');
            if (badge) {{
                badge.className = 'badge warn';
                badge.textContent = '✎ Modificado (por guardar)';
            }}
            if (issueTop) {{
                issueTop.className = 'issue ok';
                issueTop.dataset.resolved = 'true';
                const textDiv = issueTop.querySelector('.issue-text');
                if (textDiv) textDiv.innerHTML = '<strong>✓ Resuelto:</strong> Corrección aplicada (recuerda guardar cambios)';
                const link = issueTop.querySelector('.issue-link');
                if (link) link.style.display = 'none';
            }}
        }} else {{
            qBox.classList.add('question-error');
            const errAlert = qBox.querySelector('.issue.error');
            if (errAlert) errAlert.style.display = 'block';
            const badge = qBox.querySelector('.badge');
            if (badge) {{
                badge.className = 'badge error';
                badge.textContent = '❌ Requiere atención';
            }}
            if (issueTop) {{
                issueTop.className = 'issue error';
                delete issueTop.dataset.resolved;
                const textDiv = issueTop.querySelector('.issue-text');
                if (textDiv) textDiv.innerHTML = '<strong>Error</strong>: ' + (issueTop.dataset.originalMsg || '');
                const link = issueTop.querySelector('.issue-link');
                if (link) link.style.display = 'inline';
            }}
        }}

        // Recalcular errores pendientes
        const pendientes = document.querySelectorAll('#contenedor-issues .issue.error').length;
        const badgeContador = document.getElementById('badge-contador-errores');
        const badgeSinErrores = document.getElementById('badge-sin-errores');
        const numSpan = document.getElementById('num-errores');
        if (numSpan) numSpan.textContent = pendientes;
        if (pendientes === 0) {{
            if (badgeContador) badgeContador.style.display = 'none';
            if (badgeSinErrores) {{
                badgeSinErrores.style.display = 'inline-block';
                badgeSinErrores.textContent = '✓ Todas las preguntas corregidas (listo para guardar o cargar)';
            }}
        }} else {{
            if (badgeContador) badgeContador.style.display = 'inline-block';
            if (badgeSinErrores) badgeSinErrores.style.display = 'none';
        }}
    }}

    document.querySelectorAll('[data-field="answer"], [data-field="option_correct"], [data-field="prompt"]').forEach(el => {{
        el.addEventListener('change', () => actualizarEstadoPreguntaEnVivo(el.closest('.question')));
        el.addEventListener('input', () => actualizarEstadoPreguntaEnVivo(el.closest('.question')));
    }});

    document.getElementById('review-form').addEventListener('submit',()=>{{
        const data=structuredClone(initial);
        document.querySelectorAll('[data-field]').forEach(el=>{{
            const t=+el.dataset.test,s=el.dataset.section,q=el.dataset.question,f=el.dataset.field;
            if(f==='title'||f==='description'||f==='week'||f==='cut')data.tests[t][f]=val(el);
            if(f==='prompt')data.tests[t].sections[s].questions[q].prompt=val(el);
            if(f==='answer')data.tests[t].sections[s].questions[q].answer=val(el)==='true'?true:val(el)==='false'?false:null;
            if(f==='option_text')data.tests[t].sections[s].questions[q].options[+el.dataset.option].text=val(el);
            if(f==='option_correct'){{const o=data.tests[t].sections[s].questions[q].options[+el.dataset.option];o.correct=val(el);}}
            if(f==='pair_left')data.tests[t].sections[s].questions[q].pairs[+el.dataset.pair].left=val(el);
            if(f==='pair_right')data.tests[t].sections[s].questions[q].pairs[+el.dataset.pair].right=val(el);
        }});
        data.tests.forEach(t=>t.sections.forEach(sec=>sec.questions.forEach(q=>{{
            if(sec.code==='multipleChoice')q.correct_option_letters=q.options.filter(o=>o.correct).map(o=>o.letter);
        }})));
        document.getElementById('payload').value=JSON.stringify(data);
    }});

    function toggleTestCard(idx) {{
        const body = document.getElementById(`test-card-body-${{idx}}`);
        const txt = document.getElementById(`test-toggle-text-${{idx}}`);
        if (!body) return;
        if (body.style.display === 'none') {{
            body.style.display = 'block';
            if (txt) txt.textContent = 'Contraer ▲';
        }} else {{
            body.style.display = 'none';
            if (txt) txt.textContent = 'Desplegar ▼';
        }}
    }}

    function expandirTodosTests() {{
        document.querySelectorAll('.test-card-body').forEach(b => b.style.display = 'block');
        document.querySelectorAll('[id^="test-toggle-text-"]').forEach(t => t.textContent = 'Contraer ▲');
    }}

    function contraerTodosTests() {{
        document.querySelectorAll('.test-card-body').forEach(b => b.style.display = 'none');
        document.querySelectorAll('[id^="test-toggle-text-"]').forEach(t => t.textContent = 'Desplegar ▼');
    }}

    function irAPregunta(targetId, testIndex) {{
        const body = document.getElementById(`test-card-body-${{testIndex}}`);
        const txt = document.getElementById(`test-toggle-text-${{testIndex}}`);
        if (body) {{
            body.style.display = 'block';
            if (txt) txt.textContent = 'Contraer ▲';
        }}
        const qEl = document.getElementById(targetId);
        if (qEl) {{
            qEl.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
            qEl.style.transition = 'all 0.3s ease';
            qEl.style.boxShadow = '0 0 0 4px #6366f1';
            setTimeout(() => {{ qEl.style.boxShadow = ''; }}, 2200);
            const firstInput = qEl.querySelector('input, select, textarea');
            if (firstInput) firstInput.focus();
        }}
    }}

    function marcarCambiosPendientes() {{
        const btn = document.getElementById('btn-guardar-correcciones');
        if (btn && !btn.dataset.dirty) {{
            btn.dataset.dirty = 'true';
            btn.classList.add('is-dirty');
            btn.innerHTML = '💾 Guardar correcciones * (Cambios pendientes)';
        }}
    }}

    document.querySelectorAll('#review-form input, #review-form textarea, #review-form select').forEach(el => {{
        el.addEventListener('input', marcarCambiosPendientes);
        el.addEventListener('change', marcarCambiosPendientes);
    }});
    </script>'''
    return _layout("Revisión de Tests", body)


def _render_tarjeta_estado_test(info: dict | None) -> str:
    """Tarjeta de monitoreo en vivo en /tests. Solo se llama mientras el
    test esta realmente en curso (ver _obtener_estado_test_activo) - en
    cuanto termina, deja de mostrarse solo, sin intervencion del usuario."""
    if not info:
        return ""
    jid = html.escape(str(info.get("job_id") or ""))
    st = info.get("estado", "")
    titulo = html.escape(str(info.get("titulo_doc") or info.get("asignatura") or "Test Evaluativo"))
    asignatura = html.escape(str(info.get("asignatura") or ""))
    mensaje = html.escape(str(info.get("mensaje") or ""))
    pct = int(info.get("progreso_porcentaje") or 0)
    total = int(info.get("total") or 0)
    exit = int(info.get("exitosos") or 0)
    fall = int(info.get("fallidos") or 0)

    if st == "esperando_asignatura":
        clase_card = "live-card--waiting"
        badge = '<span class="badge-dinamico badge-dinamico--error">⚠️ Selección requerida</span>'
        icono = "⚠️"
    else:
        clase_card = "live-card--running"
        badge = f'<span class="badge-dinamico badge-dinamico--proceso">⏳ En proceso ({pct}%)</span>'
        icono = "🔄"

    asig_html = f' · Asignatura: <strong>{asignatura}</strong>' if asignatura else ''

    return f'''
    <section id="card-estado-test-activo" class="panel live-card {clase_card}">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:14px;margin-bottom:12px;">
            <div style="display:flex;gap:12px;align-items:center;">
                <span style="font-size:26px;">{icono}</span>
                <div>
                    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                        <h3>Cargue de Test en PRIZMA</h3>
                        <span id="card-badge-estado">{badge}</span>
                    </div>
                    <p>Documento: <strong>{titulo}</strong>{asig_html}</p>
                </div>
            </div>
            <a href="/tests/{jid}/estado-prizma" class="boton" style="text-decoration:none;padding:10px 18px;font-size:13px;">
                ⚡ Ver monitor en vivo ➔
            </a>
        </div>

        <div style="margin:12px 0 8px;">
            <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:5px;font-weight:600;">
                <span id="card-txt-avance" class="muted">{exit} de {total} tests guardados exitosamente</span>
                <span id="card-txt-pct" style="color:#5548e8;font-weight:800;">{pct}%</span>
            </div>
            <div class="progress-track">
                <div id="card-barra-progreso" class="progress-fill" style="width:{pct}%;"></div>
            </div>
        </div>

        <div style="display:flex;justify-content:space-between;align-items:center;font-size:12px;margin-top:10px;flex-wrap:wrap;gap:8px;">
            <span id="card-mensaje" class="muted" style="font-style:italic;">💬 {mensaje}</span>
            <span id="card-stats" class="muted" style="font-weight:600;">Exitosos: {exit} · Fallidos: {fall} · Total: {total}</span>
        </div>
    </section>
    '''


@router.get("/tests", response_class=HTMLResponse)
def tests_home(request: Request):
    user = _user(request)
    google_conectado = _es_google_conectado(user)
    info_activo = _obtener_estado_test_activo(user)
    tarjeta_activo = _render_tarjeta_estado_test(info_activo)

    if google_conectado:
        alerta_google = '''<div class="alert-soft alert-soft--ok">
            <span style="font-size:20px;">✓</span>
            <span style="font-size:13px;">Tu cuenta de Google está vinculada en esta sesión. Puedes ingresar enlaces privados o compartidos.</span>
        </div>'''
    else:
        alerta_google = '''<div class="alert-soft alert-soft--warn">
            <span style="font-size:20px;">⚠️</span>
            <div style="font-size:13px;line-height:1.5;">
                <strong>¿Tu documento de Google Docs es privado?</strong><br>
                <a href="/" style="font-weight:700;text-decoration:underline;">Ve al inicio e inicia sesión con tu cuenta de Google</a> para autorizar la sincronización privada. Si el documento ya tiene acceso público con el enlace, puedes continuar directamente.
            </div>
        </div>'''

    body = f'''<h1>Nuevo Test</h1>
    <p>Asistente para convertir evaluaciones académicas en tests auditados y JSON compatibles para PRIZMA de manera 100% automatizada.</p>
    <div class="pasos">
        <span class="paso">1 Entrada</span>
        <span class="paso">2 Revisión</span>
        <span class="paso">3 PRIZMA bloqueado</span>
        <span class="paso">4 Pendiente</span>
    </div>

    {tarjeta_activo}

    <div class="layout-tests">
        <section class="panel">
            <h2>Suministrar evaluación</h2>
            <p>Elige el origen del documento con las preguntas del test:</p>

            <form id="form-analizar" action="/tests/analizar" method="post" enctype="multipart/form-data">
                <div style="display:flex;gap:12px;margin-bottom:22px;flex-wrap:wrap;">
                    <label id="lbl-modo-archivo" class="mode-option mode-option--activo">
                        <input type="radio" name="modo_entrada" value="archivo" checked style="accent-color:#5548e8;">
                        <strong>📄 Por archivo</strong>
                        <span>Sube un documento Word (.docx)</span>
                    </label>
                    <label id="lbl-modo-link" class="mode-option">
                        <input type="radio" name="modo_entrada" value="link" style="accent-color:#5548e8;">
                        <strong>🔗 Por link</strong>
                        <span>Google Docs (público o privado)</span>
                    </label>
                </div>

                <div id="bloque-archivo">
                    <label class="dropzone" id="dropzone-box" for="archivo_docx">
                        <div id="dropzone-vacio">
                            <div style="font-size:36px;margin-bottom:8px;">📄</div>
                            <strong style="font-size:15px;">Arrastra tu archivo Word aquí</strong>
                            <span class="muted" style="display:block;margin-top:4px;">o haz clic para buscarlo en tu equipo (.docx)</span>
                        </div>
                        <div id="dropzone-seleccionado" style="display:none;padding:10px 0;">
                            <div style="font-size:36px;color:var(--verde);">✓</div>
                            <strong id="nombre-archivo-seleccionado" style="font-size:15px;display:block;margin:6px 0 4px;"></strong>
                            <span id="peso-archivo-seleccionado" class="badge ok" style="margin-bottom:8px;"></span>
                            <div style="margin-top:8px;">
                                <span style="font-size:12px;color:var(--morado);text-decoration:underline;">Haz clic aquí para seleccionar otro archivo</span>
                            </div>
                        </div>
                        <input id="archivo_docx" name="archivo_docx" type="file" accept=".docx" hidden>
                    </label>
                </div>

                <div id="bloque-link" style="display:none;">
                    <label for="google_doc_url">Enlace del Google Doc</label>
                    <input id="google_doc_url" name="google_doc_url" type="url" placeholder="https://docs.google.com/document/d/.../edit">
                    {alerta_google}
                </div>

                <div class="acciones">
                    <button type="submit" id="btn-analizar">🚀 Analizar documento</button>
                </div>
            </form>
        </section>

        <aside class="panel-ayuda">
            <h3>¿Cómo funciona?</h3>
            <div class="paso-ayuda"><b>1</b><span>Elige <strong>Por archivo</strong> (.docx) o <strong>Por link</strong> (Google Docs).</span></div>
            <div class="paso-ayuda"><b>2</b><span>El sistema extraerá automáticamente el <strong>Título</strong>, <strong>Semana</strong>, <strong>Corte</strong> y preguntas.</span></div>
            <div class="paso-ayuda"><b>3</b><span>Audita y ajusta las preguntas en la pantalla de revisión con semáforo pedagógico.</span></div>
            <div class="paso-ayuda"><b>4</b><span>Descarga el respaldo en JSON o déjalos preparados para la inyección desatendida en PRIZMA.</span></div>

            <div class="ayuda-separador"></div>

            <h3 class="titulo-consejos">Consejos</h3>
            <p class="consejo">✓ Cada test debe incluir su Semana, Corte y Título correspondiente.</p>
            <p class="consejo">✓ Las preguntas cerradas deben incluir su clave (ej: <code>[R: Verdadero]</code> o <code>[R: b]</code>).</p>
            <p class="consejo">✓ Si el documento contiene varios tests evaluativos, se dividirán automáticamente en tarjetas independientes.</p>
        </aside>
    </div>

    <script>
    const radioArchivo = document.querySelector('input[name="modo_entrada"][value="archivo"]');
    const radioLink = document.querySelector('input[name="modo_entrada"][value="link"]');
    const lblArchivo = document.getElementById('lbl-modo-archivo');
    const lblLink = document.getElementById('lbl-modo-link');
    const bloqueArchivo = document.getElementById('bloque-archivo');
    const bloqueLink = document.getElementById('bloque-link');
    const inputFile = document.getElementById('archivo_docx');
    const dropzoneBox = document.getElementById('dropzone-box');
    const dropVacio = document.getElementById('dropzone-vacio');
    const dropSeleccionado = document.getElementById('dropzone-seleccionado');
    const nombreArchivo = document.getElementById('nombre-archivo-seleccionado');
    const pesoArchivo = document.getElementById('peso-archivo-seleccionado');
    const btnAnalizar = document.getElementById('btn-analizar');

    function actualizarModo() {{
        const esArchivo = radioArchivo.checked;
        bloqueArchivo.style.display = esArchivo ? '' : 'none';
        bloqueLink.style.display = esArchivo ? 'none' : '';
        lblArchivo.classList.toggle('mode-option--activo', esArchivo);
        lblLink.classList.toggle('mode-option--activo', !esArchivo);
    }}

    radioArchivo.addEventListener('change', actualizarModo);
    radioLink.addEventListener('change', actualizarModo);

    function mostrarArchivo(file) {{
        if (!file) return;
        nombreArchivo.textContent = file.name;
        const kb = (file.size / 1024).toFixed(1);
        const mb = (file.size / (1024 * 1024)).toFixed(2);
        pesoArchivo.textContent = file.size > 1024 * 1024 ? `${{mb}} MB` : `${{kb}} KB`;
        dropVacio.style.display = 'none';
        dropSeleccionado.style.display = 'block';
    }}

    inputFile.addEventListener('change', (e) => {{
        if (e.target.files && e.target.files.length > 0) {{
            mostrarArchivo(e.target.files[0]);
        }}
    }});

    ['dragenter', 'dragover'].forEach(eventName => {{
        dropzoneBox.addEventListener(eventName, (e) => {{
            e.preventDefault();
            e.stopPropagation();
            dropzoneBox.classList.add('dragover');
        }}, false);
    }});

    ['dragleave', 'drop'].forEach(eventName => {{
        dropzoneBox.addEventListener(eventName, (e) => {{
            e.preventDefault();
            e.stopPropagation();
            dropzoneBox.classList.remove('dragover');
        }}, false);
    }});

    dropzoneBox.addEventListener('drop', (e) => {{
        const dt = e.dataTransfer;
        if (dt && dt.files && dt.files.length > 0) {{
            inputFile.files = dt.files;
            mostrarArchivo(dt.files[0]);
        }}
    }});

    document.getElementById('form-analizar').addEventListener('submit', () => {{
        btnAnalizar.disabled = true;
        btnAnalizar.textContent = '⏳ Analizando documento...';
    }});
    </script>'''
    if info_activo and info_activo.get("estado") in ["en_ejecucion", "esperando_asignatura"]:
        body += '''<script>
        (function() {
            function refrescarEstadoCard() {
                fetch('/tests/estado-activo/json')
                    .then(r => r.json())
                    .then(data => {
                        if (!data || !data.activo) {
                            // El test ya no esta activo (termino mientras
                            // mirabamos): recargar para que la tarjeta
                            // desaparezca sola, sin quedar congelada.
                            setTimeout(() => window.location.reload(), 1200);
                            return;
                        }
                        const elPct = document.getElementById('card-txt-pct');
                        const elBar = document.getElementById('card-barra-progreso');
                        const elMsg = document.getElementById('card-mensaje');
                        const elAvance = document.getElementById('card-txt-avance');
                        const elStats = document.getElementById('card-stats');
                        if (elPct) elPct.textContent = (data.progreso_porcentaje || 0) + '%';
                        if (elBar) elBar.style.width = (data.progreso_porcentaje || 0) + '%';
                        if (elMsg && data.mensaje) elMsg.textContent = '💬 ' + data.mensaje;
                        if (elAvance) elAvance.textContent = `${data.exitosos || 0} de ${data.total || 0} tests guardados exitosamente`;
                        if (elStats) elStats.textContent = `Exitosos: ${data.exitosos || 0} · Fallidos: ${data.fallidos || 0} · Total: ${data.total || 0}`;
                        if (data.estado !== 'en_ejecucion' && data.estado !== 'esperando_asignatura') {
                            setTimeout(() => window.location.reload(), 1500);
                        } else {
                            setTimeout(refrescarEstadoCard, 2000);
                        }
                    })
                    .catch(() => setTimeout(refrescarEstadoCard, 4000));
            }
            setTimeout(refrescarEstadoCard, 2000);
        })();
        </script>'''
    return HTMLResponse(_layout("Nuevo Test", body))


@router.post("/tests/analizar")
async def analyze_tests(
    request: Request,
    archivo_docx: UploadFile | None = File(default=None),
    google_doc_url: str | None = Form(default=None),
    modo_entrada: str = Form(default="archivo"),
):
    content = b""
    source_name = ""
    modo_entrada = (modo_entrada or "archivo").strip().lower()

    if modo_entrada == "archivo" and archivo_docx and archivo_docx.filename:
        if not archivo_docx.filename.lower().endswith(".docx"):
            return HTMLResponse(
                _layout("Test", '<section class="panel"><div class="issue error">El archivo debe tener formato Microsoft Word (.docx).</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
                status_code=400,
            )
        content = await archivo_docx.read()
        source_name = archivo_docx.filename
    elif modo_entrada == "link" or (google_doc_url and google_doc_url.strip()):
        url_doc = str(google_doc_url or "").strip()
        if not url_doc:
            return HTMLResponse(
                _layout("Test", '<section class="panel"><div class="issue error">Ingresa el enlace del Google Doc.</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
                status_code=400,
            )
        loader = getattr(request.app.state, "test_google_loader", None)
        if not loader:
            return HTMLResponse(
                _layout("Test", '<section class="panel"><div class="issue error">La conexión con Google Drive no está inicializada en el servidor.</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
                status_code=400,
            )
        try:
            content, source_name = loader(url_doc, _user(request))
        except Exception as exc:
            return HTMLResponse(
                _layout("Test", f'<section class="panel"><div class="issue error">No se pudo descargar el Google Doc: {html.escape(str(exc))}</div><div class="acciones"><a class="boton" href="/tests">Volver a intentar</a></div></section>'),
                status_code=400,
            )
    elif archivo_docx and archivo_docx.filename:
        if not archivo_docx.filename.lower().endswith(".docx"):
            return HTMLResponse(
                _layout("Test", '<section class="panel"><div class="issue error">El archivo debe tener formato Microsoft Word (.docx).</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
                status_code=400,
            )
        content = await archivo_docx.read()
        source_name = archivo_docx.filename
    else:
        return HTMLResponse(
            _layout("Test", '<section class="panel"><div class="issue error">Selecciona un archivo Word (.docx) o introduce un enlace de Google Docs.</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
            status_code=400,
        )

    if not content:
        return HTMLResponse(
            _layout("Test", '<section class="panel"><div class="issue error">El documento seleccionado está vacío o no contiene datos procesables.</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
            status_code=400,
        )

    try:
        document = parse_docx_bytes(content, source_name=source_name)
    except Exception as exc:
        return HTMLResponse(
            _layout("Test", f'<section class="panel"><div class="issue error">Error al procesar el archivo Word: {html.escape(str(exc))}</div><div class="acciones"><a class="boton" href="/tests">Volver</a></div></section>'),
            status_code=400,
        )

    job_id = uuid.uuid4().hex
    with _JOBS_LOCK:
        _JOBS[job_id] = {
            "owner": _user(request),
            "document": document,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
    return RedirectResponse(f"/tests/{job_id}/revision", status_code=303)


@router.get("/tests/{job_id}/revision", response_class=HTMLResponse)
def review_tests(request: Request, job_id: str):
    job = _job_for(request, job_id)
    if not job:
        return HTMLResponse(_layout("Test", '<section class="panel"><div class="issue error">Análisis no encontrado.</div><a class="boton" href="/tests">Volver</a></section>'), status_code=404)
    return HTMLResponse(_render_review(job_id, job))


@router.post("/tests/{job_id}/guardar")
def save_review(request: Request, job_id: str, payload: str = Form(...)):
    job = _job_for(request, job_id)
    if not job:
        return HTMLResponse(_layout("Test", '<section class="panel"><div class="issue error">Análisis no encontrado.</div></section>'), status_code=404)
    try:
        data = json.loads(payload)
        document = document_from_dict(data)
    except (TypeError, ValueError, json.JSONDecodeError) as exc:
        return HTMLResponse(_render_review(job_id, job, f"No se pudieron guardar las correcciones: {exc}"), status_code=400)
    job["document"] = document
    return HTMLResponse(_render_review(job_id, job, "Correcciones guardadas y validadas."))


@router.get("/tests/{job_id}/descargar")
def download_tests(request: Request, job_id: str, indice: int | None = None):
    job = _job_for(request, job_id)
    if not job:
        return JSONResponse({"error": "Análisis no encontrado"}, status_code=404)
    document: TestDocument = job["document"]
    if indice is not None:
        if indice < 0 or indice >= len(document.tests):
            return JSONResponse({"error": "Test no encontrado"}, status_code=404)
        payload = json.dumps(serialize_test(document.tests[indice]), ensure_ascii=False, indent=2).encode("utf-8")
        return Response(payload, media_type="application/json", headers={"Content-Disposition": f'attachment; filename="test_{indice + 1}.json"'})
    if len(document.tests) == 1:
        payload = json.dumps(serialize_test(document.tests[0]), ensure_ascii=False, indent=2).encode("utf-8")
        return Response(payload, media_type="application/json", headers={"Content-Disposition": 'attachment; filename="test.json"'})
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        for index, test in enumerate(document.tests, start=1):
            bundle.writestr(f"{index:02d}_{test.title[:60] or 'test'}.json", json.dumps(serialize_test(test), ensure_ascii=False, indent=2))
    return Response(archive.getvalue(), media_type="application/zip", headers={"Content-Disposition": 'attachment; filename="tests_json.zip"'})


def simplificar_mensaje_error(raw_error: str | Exception | None) -> str:
    """Traduce excepciones o trazas técnicas de Playwright en explicaciones limpias y ejecutivas."""
    if not raw_error:
        return "Ocurrió un error inesperado durante la ejecución en PRIZMA."
    msg = str(raw_error).strip()

    if "se encuentra deshabilitado" in msg or "continúa deshabilitado" in msg:
        return "El botón 'Crear' en PRIZMA se encuentra deshabilitado. Verifica que todos los campos requeridos estén completos."
    if "element is not enabled" in msg.lower():
        return "No se pudo presionar el botón 'Crear' en PRIZMA porque se encuentra deshabilitado. Verifica que todos los campos requeridos estén completos."
    if "No se encontró el botón 'Crear'" in msg:
        return "No se localizó el botón 'Crear' al lado de 'Cancelar' en el formulario activo de PRIZMA."
    if "No se encontró el selector 'Facultad'" in msg:
        return "No se encontró el selector 'Facultad' en el formulario de PRIZMA tras acceder a Actividades -> Tests -> Crear."
    if "Campo obligatorio o inválido" in msg or "Campo requerido o inválido" in msg:
        return msg.split("\n")[0].strip()
    if "Error de PRIZMA al guardar" in msg:
        return msg.split("\n")[0].strip()
    if "inicio-sesion" in msg.lower() or "credenciales" in msg.lower():
        return "Error al iniciar sesión en PRIZMA. Verifica tu usuario, contraseña y conexión a la plataforma."
    if "Importar" in msg and ("timeout" in msg.lower() or "waiting for locator" in msg.lower()):
        return "No se localizó el botón 'Importar' dentro del tiempo de espera en el formulario de PRIZMA."

    if "Call log:" in msg:
        msg = msg.split("Call log:")[0].strip()

    lineas = [l.strip() for l in msg.splitlines() if l.strip()]
    if lineas:
        primera = lineas[0]
        primera = re.sub(r"^(Error durante la automatización:\s*)+", "", primera)
        primera = re.sub(r"^(Locator\.[a-zA-Z_]+:\s*)+", "", primera)
        if "Timeout" in primera and "exceeded" in primera:
            return "Tiempo de espera agotado al interactuar con la plataforma PRIZMA. Revisa la conexión o el estado de la página."
        return primera[:250].strip()

    return "Error durante la interacción con la plataforma PRIZMA."


def ejecutar_test_job(
    job_id: str,
    usuario_prizma: str,
    clave_prizma: str,
    user: str = "",
    simulacion: bool = False,
    headless: bool = True,
    cascada_override: dict | None = None,
    on_test_callback=None,
    on_log_callback=None,
    cancel_checker=None,
    ruta_reporte: str | None = None,
    registrar_historial: bool = False,
) -> dict:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        return {"ok": False, "error": "Análisis de test no encontrado en el servidor."}

    document: TestDocument = job["document"]
    cascada = dict(cascada_override or job.get("cascada_datos") or {})

    total_tests = len(document.tests)
    with _AUTOMATION_LOCK:
        detalle_tests = [
            {
                "numero": idx,
                "titulo": t.title,
                "preguntas": sum(len(sec.questions) for sec in t.sections),
                "estado": "pendiente",
                "mensaje": "En cola de espera...",
            }
            for idx, t in enumerate(document.tests, start=1)
        ]
        auto_state = {
            "job_id": job_id,
            "owner": user or job.get("owner", ""),
            "simulacion": simulacion,
            "cancelado": False,
            "estado": "en_ejecucion",
            "mensaje": (
                "Iniciando motor de automatización Playwright en modo simulación (sin guardar)..."
                if simulacion
                else "Iniciando motor de automatización Playwright en modo creación real..."
            ),
            "total": total_tests,
            "procesados": 0,
            "exitosos": 0,
            "pendientes": total_tests,
            "fallidos": 0,
            "progreso_porcentaje": 0,
            "detalle_tests": detalle_tests,
            "logs": [],
            "captura": None,
            "error": "",
            "esperando_asignatura": False,
            "opciones_asignatura": [],
            "asignatura_buscada": "",
            "iniciado_en": datetime.now(timezone.utc).isoformat(),
            "finalizado_en": None,
            "asignatura": cascada.get("asignatura", ""),
            "titulo_doc": getattr(document, "source_name", None) or (document.tests[0].title if document.tests else "Test Evaluativo"),
        }
        _AUTOMATION_JOBS[job_id] = auto_state
        sync_obj = {
            "event": threading.Event(),
            "seleccion": {"asignatura": None},
        }
        _AUTOMATION_SYNC[job_id] = sync_obj

    import time as time_mod

    def on_log(nivel, msg, snap=None):
        with _AUTOMATION_LOCK:
            auto_state["mensaje"] = msg
            auto_state["logs"].append({"tiempo": time_mod.strftime("%H:%M:%S"), "nivel": nivel, "mensaje": msg})
            if snap:
                auto_state["captura"] = snap
        if on_log_callback:
            try:
                on_log_callback(nivel, msg, snap)
            except Exception:
                pass

    def pedir_seleccion_asignatura(opciones: list[str], buscada: str) -> str:
        with _AUTOMATION_LOCK:
            auto_state["esperando_asignatura"] = True
            auto_state["opciones_asignatura"] = opciones
            auto_state["asignatura_buscada"] = buscada
            msg_espera = (
                f"La asignatura '{buscada}' no coincide exactamente en PRIZMA. "
                f"Por favor selecciona una de las opciones en pantalla para continuar."
            )
            auto_state["mensaje"] = msg_espera
            auto_state["logs"].append({
                "tiempo": time_mod.strftime("%H:%M:%S"),
                "nivel": "warn",
                "mensaje": msg_espera,
            })
        sync_obj["event"].clear()
        notificado = sync_obj["event"].wait(timeout=300)
        with _AUTOMATION_LOCK:
            auto_state["esperando_asignatura"] = False
            auto_state["opciones_asignatura"] = []
        if not notificado or not sync_obj["seleccion"].get("asignatura"):
            raise ValueError("Tiempo de espera agotado (5 min) para seleccionar la asignatura en PRIZMA.")

        asignatura_confirmada = str(sync_obj["seleccion"]["asignatura"] or "").strip()
        with _AUTOMATION_LOCK:
            cascada["asignatura"] = asignatura_confirmada
            auto_state["asignatura"] = asignatura_confirmada
        with _JOBS_LOCK:
            job["cascada_datos"] = dict(cascada)
        sync_obj["seleccion"]["asignatura"] = None
        return asignatura_confirmada

    def _fila_test_reporte(indice, item, resultado, error_msg=""):
        if not ruta_reporte:
            return
        t = document.tests[indice - 1] if 0 < indice <= len(document.tests) else None
        actividad = {
            "fila_excel": f"Test {indice}",
            "programa": cascada.get("programa") or "",
            "curso": cascada.get("asignatura") or "",
            "semana": (getattr(t, "week", "") or cascada.get("nivel") or "") if t else "",
            "unidad": (getattr(t, "cut", "") or "") if t else "",
            "nombre": item.get("titulo") or f"Test {indice}",
            "categoria_prizma": "Test Evaluativo",
            "tipo_archivo": "Test",
        }
        try:
            motor_prizma.guardar_resultado(ruta_reporte, actividad, resultado, error_msg)
        except Exception:
            pass

    def on_test_progreso(indice: int, estado_test: str, msg_test: str, err: str = ""):
        with _AUTOMATION_LOCK:
            if 1 <= indice <= len(auto_state["detalle_tests"]):
                item = auto_state["detalle_tests"][indice - 1]
                item["estado"] = estado_test
                item["mensaje"] = msg_test
            if estado_test == "exitoso":
                auto_state["exitosos"] += 1
                auto_state["procesados"] += 1
                auto_state["pendientes"] = max(0, auto_state["total"] - auto_state["procesados"])
                _fila_test_reporte(indice, item, "Cargado")
            elif estado_test == "error":
                auto_state["fallidos"] += 1
                auto_state["procesados"] += 1
                auto_state["pendientes"] = max(0, auto_state["total"] - auto_state["procesados"])
                _fila_test_reporte(indice, item, "Error", err or msg_test)
            elif estado_test == "cancelado":
                auto_state["pendientes"] = max(0, auto_state["total"] - auto_state["procesados"])
                _fila_test_reporte(indice, item, "Cancelado")

            total_t = max(1, auto_state["total"])
            auto_state["progreso_porcentaje"] = int((auto_state["procesados"] / total_t) * 100)
        if on_test_callback:
            try:
                on_test_callback(indice, estado_test, msg_test, err)
            except Exception:
                pass

    def _is_cancelled():
        if auto_state.get("cancelado"):
            return True
        if cancel_checker and cancel_checker():
            return True
        return False

    def _registrar_historial_si_aplica():
        if not registrar_historial or not ruta_reporte:
            return
        estado_final = "finalizado" if auto_state["estado"] == "completado" else auto_state["estado"]
        try:
            from main import _registrar_historial_test
            _registrar_historial_test(job_id, ruta_reporte, cascada, estado_final, auto_state.get("owner", ""))
        except Exception:
            pass

    try:
        automator = PrizmaTestAutomator(
            base_url="https://admin.prizma.site",
            headless=headless,
            log_callback=on_log,
        )
        res = automator.ejecutar_cargue(
            usuario_prizma=usuario_prizma,
            clave_prizma=clave_prizma,
            tests=document.tests,
            cascada=cascada,
            solo_simulacion=simulacion,
            test_callback=on_test_progreso,
            cancel_checker=_is_cancelled,
            solicitar_asignatura_callback=pedir_seleccion_asignatura,
        )
        with _AUTOMATION_LOCK:
            auto_state["finalizado_en"] = datetime.now(timezone.utc).isoformat()
            if res.get("detenido") or auto_state.get("cancelado"):
                auto_state["estado"] = "detenido"
                auto_state["mensaje"] = f"⏹ Automatización detenida por el usuario. Se completaron {auto_state.get('exitosos', 0)} de {auto_state.get('total', 0)} tests."
                auto_state["pendientes"] = max(0, auto_state["total"] - auto_state["procesados"])
            elif res.get("ok"):
                auto_state["estado"] = "completado"
                auto_state["progreso_porcentaje"] = 100
                auto_state["pendientes"] = 0
                if simulacion:
                    auto_state["mensaje"] = "🛡 Simulación exitosa: todos los tests fueron validados sin guardar en PRIZMA."
                else:
                    auto_state["mensaje"] = "¡Todos los tests fueron creados y guardados exitosamente en PRIZMA!"
                auto_state["procesados"] = res.get("procesados", len(document.tests))
            else:
                auto_state["estado"] = "error"
                raw_err = res.get("error", "Error desconocido en la automatización.")
                auto_state["error"] = simplificar_mensaje_error(raw_err)
                auto_state["mensaje"] = f"Fallo en la ejecución: {auto_state['error']}"
        _registrar_historial_si_aplica()
        return res
    except Exception as e:
        with _AUTOMATION_LOCK:
            auto_state["finalizado_en"] = datetime.now(timezone.utc).isoformat()
            if auto_state.get("cancelado"):
                auto_state["estado"] = "detenido"
                auto_state["mensaje"] = f"⏹ Automatización detenida por el usuario. Se completaron {auto_state.get('exitosos', 0)} de {auto_state.get('total', 0)} tests."
            else:
                auto_state["estado"] = "error"
                auto_state["error"] = simplificar_mensaje_error(e)
                auto_state["mensaje"] = f"Excepción en ejecución: {auto_state['error']}"
        _registrar_historial_si_aplica()
        return {"ok": False, "error": str(e)}


@router.post("/tests/{job_id}/ejecutar-prizma")
async def ejecutar_prizma(
    request: Request,
    job_id: str,
    prizma_usuario: str = Form(...),
    prizma_clave: str = Form(...),
    recordar_credenciales: str | None = Form(default=None),
    ver_navegador: str | None = Form(default=None),
    solo_simulacion: str | None = Form(default=None),
    cascada_facultad: str = Form(default=""),
    cascada_programa: str = Form(default=""),
    cascada_pensum: str = Form(default=""),
    cascada_nivel: str = Form(default=""),
    cascada_asignatura: str = Form(default=""),
    cascada_semana: str = Form(default=""),
    cascada_corte: str = Form(default=""),
):
    job = _job_for(request, job_id)
    if not job:
        return HTMLResponse(_layout("Test", '<section class="panel"><div class="issue error">Análisis no encontrado.</div></section>'), status_code=404)

    user = _user(request)
    recordar = bool(recordar_credenciales)
    simulacion = bool(solo_simulacion)
    _guardar_credenciales_prizma(user, {"usuario": prizma_usuario.strip() if recordar else "", "recordar": recordar})

    document: TestDocument = job["document"]
    issues = validate_document(document)
    errores = [i for i in issues if i.severity == Severity.ERROR]
    if errores:
        return HTMLResponse(_render_review(job_id, job, "Corrige los errores pendientes antes de cargar el test a PRIZMA."), status_code=400)

    headless = not bool(ver_navegador)
    cascada = {
        "facultad": cascada_facultad.strip(),
        "programa": cascada_programa.strip(),
        "pensum": cascada_pensum.strip(),
        "nivel": cascada_nivel.strip(),
        "asignatura": cascada_asignatura.strip(),
        "semana": cascada_semana.strip(),
        "corte": cascada_corte.strip(),
    }

    ruta_reporte_test = None
    if not simulacion:
        from main import _ruta_reporte_test
        ruta_reporte_test = _ruta_reporte_test(cascada, job_id)

    threading.Thread(
        target=ejecutar_test_job,
        kwargs={
            "job_id": job_id,
            "usuario_prizma": prizma_usuario.strip(),
            "clave_prizma": prizma_clave.strip(),
            "user": user,
            "simulacion": simulacion,
            "headless": headless,
            "cascada_override": cascada,
            "ruta_reporte": ruta_reporte_test,
            "registrar_historial": bool(ruta_reporte_test),
        },
        daemon=True,
    ).start()

    return RedirectResponse(f"/tests/{job_id}/estado-prizma", status_code=303)


@router.get("/tests/{job_id}/estado-prizma", response_class=HTMLResponse)
def status_prizma(request: Request, job_id: str):
    job = _job_for(request, job_id)
    if not job:
        return HTMLResponse(_layout("Test", '<section class="panel"><div class="issue error">Análisis no encontrado.</div></section>'), status_code=404)

    with _AUTOMATION_LOCK:
        auto_state = _AUTOMATION_JOBS.get(job_id)

    if not auto_state:
        return RedirectResponse(f"/tests/{job_id}/revision", status_code=303)

    es_sim = auto_state.get("simulacion", True)
    texto_completado = "✓ Simulación exitosa (Validado sin guardar)" if es_sim else "✓ Guardado completado en PRIZMA"

    rows_tests = []
    for item in auto_state.get("detalle_tests", []):
        st = item.get("estado", "pendiente")
        badge_mod = "ok" if st == "exitoso" else "error" if st == "error" else "proceso" if st == "en_proceso" else "pendiente"
        badge_txt = (
            "✓ Guardado"
            if (st == "exitoso" and not es_sim)
            else "✓ Validado"
            if st == "exitoso"
            else "❌ Error"
            if st == "error"
            else "⏳ En proceso..."
            if st == "en_proceso"
            else "⚪ En espera"
        )
        rows_tests.append(
            f'<tr id="row-test-{item["numero"]}" style="border-bottom:1px solid var(--borde);">'
            f'<td style="padding:10px 8px;font-weight:700;">{item["numero"]}</td>'
            f'<td style="padding:10px 8px;"><strong>{html.escape(item["titulo"])}</strong></td>'
            f'<td style="padding:10px 8px;"><span class="chip-morado">{item["preguntas"]} preg.</span></td>'
            f'<td style="padding:10px 8px;"><span class="badge-dinamico badge-dinamico--{badge_mod}" id="badge-test-{item["numero"]}">{badge_txt}</span></td>'
            f'<td class="muted" style="padding:10px 8px;font-size:12px;" id="msg-test-{item["numero"]}">{html.escape(item["mensaje"])}</td>'
            f'</tr>'
        )
    filas_html = "".join(rows_tests)

    body = f'''<h1>{'Monitoreo de Simulación en PRIZMA' if es_sim else 'Monitoreo de Creación en PRIZMA'}</h1>
    <p>Visualización y auditoría en tiempo real de cada test en <strong>admin.prizma.site</strong>.</p>
    <div class="pasos">
        <span class="paso">1 Entrada ✓</span>
        <span class="paso">2 Revisión ✓</span>
        <span class="paso">3 PRIZMA Activo 🚀</span>
        <span class="paso" id="paso-final">{'4 Simulación (Sin guardar)' if es_sim else '4 Creación en Plataforma'}</span>
    </div>

    <!-- BANNER DE ERROR SUPERIOR DESTACADO -->
    <div id="banner-error-superior" class="status-banner status-banner--error" style="display:{'flex' if auto_state.get('estado') == 'error' else 'none'};">
        <div class="status-icon">✕</div>
        <div style="flex:1;">
            <h3>La ejecución se detuvo por un error</h3>
            <p id="texto-error-superior">{html.escape(auto_state.get('error') or auto_state.get('mensaje') or '')}</p>
            <div style="display:flex;gap:12px;flex-wrap:wrap;">
                <a class="boton secundario" href="/tests/{job_id}/revision" style="padding:9px 16px;font-size:13px;">← Volver a Revisión para corregir parámetros</a>
            </div>
        </div>
    </div>

    <!-- BANNER DE ÉXITO SUPERIOR -->
    <div id="banner-exito-superior" class="status-banner status-banner--ok" style="display:{'flex' if auto_state.get('estado') == 'completado' else 'none'};">
        <div class="status-icon">✓</div>
        <div style="flex:1;">
            <h3>{'Simulación validada exitosamente' if es_sim else '¡Tests creados y guardados con éxito en PRIZMA!'}</h3>
            <p id="texto-exito-superior" style="margin:0;">{html.escape(auto_state.get('mensaje', ''))}</p>
        </div>
    </div>

    <!-- BANNER DE DETENCIÓN SUPERIOR -->
    <div id="banner-detenido-superior" class="status-banner status-banner--warn" style="display:{'flex' if auto_state.get('estado') == 'detenido' else 'none'};">
        <div class="status-icon">⏹</div>
        <div style="flex:1;">
            <h3>Automatización detenida por el usuario</h3>
            <p id="texto-detenido-superior">{html.escape(auto_state.get('mensaje', ''))}</p>
            <div style="display:flex;gap:12px;flex-wrap:wrap;">
                <a class="boton secundario" href="/tests/{job_id}/revision" style="padding:9px 16px;font-size:13px;">← Volver a Revisión</a>
            </div>
        </div>
    </div>

    <!-- PANEL INTERACTIVO DE SELECCIÓN DE ASIGNATURA -->
    <div id="panel-seleccion-asignatura" class="status-banner status-banner--warn" style="display:{'flex' if auto_state.get('esperando_asignatura') else 'none'};flex-direction:column;">
        <div style="display:flex;gap:14px;align-items:flex-start;width:100%;">
            <div class="status-icon">⚠️</div>
            <div style="flex:1;">
                <h3>Selección de Asignatura Requerida</h3>
                <p id="txt-asignatura-buscada" style="margin:0;">La asignatura ingresada no coincide exactamente con las opciones registradas en PRIZMA. Selecciona la opción correcta para que el robot continúe:</p>
            </div>
        </div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin-top:14px;width:100%;">
            <select id="sel-asignatura-opciones" style="flex:1;min-width:280px;">
                <option value="">-- Seleccionar asignatura de PRIZMA --</option>
            </select>
            <button id="btn-confirmar-asignatura" type="button" style="background:#d97706;">
                Confirmar Asignatura y Continuar ➔
            </button>
        </div>
        <div id="msg-confirmacion-asignatura" style="margin-top:10px;font-size:13px;font-weight:600;display:none;"></div>
    </div>

    <!-- TARJETAS KPI DE PROGRESO -->
    <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(180px, 1fr));gap:14px;margin-bottom:20px;">
        <div class="kpi-card">
            <div class="kpi-label">Total Tests</div>
            <div id="kpi-total" class="kpi-valor">{auto_state.get('total', 0)}</div>
        </div>
        <div class="kpi-card kpi-card--ok">
            <div class="kpi-label">{'✓ Validados' if es_sim else '✅ Guardados'}</div>
            <div id="kpi-exitosos" class="kpi-valor">{auto_state.get('exitosos', 0)}</div>
        </div>
        <div class="kpi-card kpi-card--warn">
            <div class="kpi-label">⏳ Faltantes / En Cola</div>
            <div id="kpi-pendientes" class="kpi-valor">{auto_state.get('pendientes', 0)}</div>
        </div>
        <div class="kpi-card kpi-card--error">
            <div class="kpi-label">❌ Con Errores</div>
            <div id="kpi-fallidos" class="kpi-valor">{auto_state.get('fallidos', 0)}</div>
        </div>
    </div>

    <!-- BARRA DE PROGRESO VISUAL -->
    <div class="panel">
        <div style="display:flex;justify-content:space-between;margin-bottom:8px;font-size:13px;font-weight:700;">
            <span id="progreso-texto-avance">Avance general ({auto_state.get('procesados', 0)} de {auto_state.get('total', 0)} tests)</span>
            <span id="progreso-porcentaje-texto">{auto_state.get('progreso_porcentaje', 0)}%</span>
        </div>
        <div class="progress-track">
            <div id="progreso-barra-fill" class="progress-fill" style="width:{auto_state.get('progreso_porcentaje', 0)}%;"></div>
        </div>
    </div>

    <!-- TABLA DE DETALLE POR TEST -->
    <div class="panel" style="overflow-x:auto;">
        <h3 style="margin:0 0 12px;font-size:15px;">Detalle de Tests en Proceso:</h3>
        <table style="width:100%;border-collapse:collapse;font-size:13px;text-align:left;">
            <thead>
                <tr style="border-bottom:2px solid var(--borde);">
                    <th style="padding:10px 8px;width:50px;">#</th>
                    <th style="padding:10px 8px;">Título del Test</th>
                    <th style="padding:10px 8px;width:110px;">Preguntas</th>
                    <th style="padding:10px 8px;width:170px;">Estado</th>
                    <th style="padding:10px 8px;">Observación</th>
                </tr>
            </thead>
            <tbody id="tabla-tests-body">
                {filas_html}
            </tbody>
        </table>
    </div>

    <section class="panel">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:10px;">
            <span class="badge {'ok' if auto_state['estado'] == 'completado' else 'error' if auto_state['estado'] == 'error' else 'warn'}" id="estado-badge" style="font-size:14px;padding:8px 16px;">
                {texto_completado if auto_state['estado'] == 'completado' else '❌ Error en la ejecución' if auto_state['estado'] == 'error' else '⏳ En ejecución...'}
            </span>
            <span class="muted" id="estado-resumen">{auto_state.get('procesados', 0)} de {auto_state.get('total', 0)} test(s) procesados</span>
        </div>
        <h3 id="estado-mensaje" style="margin:0 0 16px;">{html.escape(auto_state.get('mensaje', ''))}</h3>

        <div class="acciones" style="margin-top:24px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
            <button id="btn-detener-auto" type="button" class="boton secundario peligro" onclick="detenerAutomatizacion()" style="display:{'inline-flex' if auto_state.get('estado') == 'en_ejecucion' else 'none'};align-items:center;gap:6px;cursor:pointer;">⏹ Detener automatización</button>
            <a class="boton secundario" href="/tests/{job_id}/revision">← Volver a Revisión</a>
            <a class="boton secundario" href="/tests">＋ Nuevo análisis</a>
            <a class="boton" href="/" style="background:var(--verde);">✓ Ir a Inicio</a>
        </div>
    </section>

    <script>
    const jobId = "{job_id}";
    const badge = document.getElementById("estado-badge");
    const msg = document.getElementById("estado-mensaje");
    const resumen = document.getElementById("estado-resumen");

    let terminado = {"true" if auto_state['estado'] in ('completado', 'error', 'detenido') else "false"};

    function detenerAutomatizacion() {{
        if (!confirm("¿Confirmas que deseas detener la subida de los tests pendientes? Los tests que ya hayan sido guardados se conservarán.")) return;
        const btn = document.getElementById("btn-detener-auto");
        if (btn) {{
            btn.disabled = true;
            btn.textContent = "⏳ Deteniendo...";
        }}
        fetch(`/tests/${{jobId}}/cancelar-prizma`, {{ method: "POST" }})
            .then(r => r.json())
            .then(data => {{
                if (data.ok) {{
                    msg.textContent = "Deteniendo automatización a solicitud del usuario...";
                }}
            }})
            .catch(() => {{}});
    }}

    function actualizarDetalleTests(detalle, esSimulacion) {{
        if (!detalle || !Array.isArray(detalle)) return;
        detalle.forEach(item => {{
            const b = document.getElementById("badge-test-" + item.numero);
            const m = document.getElementById("msg-test-" + item.numero);
            if (m && item.mensaje) m.textContent = item.mensaje;
            if (b) {{
                if (item.estado === "exitoso") {{
                    b.className = "badge-dinamico badge-dinamico--ok";
                    b.textContent = esSimulacion ? "✓ Validado" : "✓ Guardado";
                }} else if (item.estado === "error") {{
                    b.className = "badge-dinamico badge-dinamico--error";
                    b.textContent = "❌ Error";
                }} else if (item.estado === "cancelado") {{
                    b.className = "badge-dinamico badge-dinamico--detenido";
                    b.textContent = "⏹ Detenido";
                }} else if (item.estado === "en_proceso") {{
                    b.className = "badge-dinamico badge-dinamico--proceso";
                    b.textContent = "⏳ En proceso...";
                }} else {{
                    b.className = "badge-dinamico badge-dinamico--pendiente";
                    b.textContent = "⚪ En espera";
                }}
            }}
        }});
    }}

    function pollEstado() {{
        if (terminado) return;
        fetch(`/tests/${{jobId}}/estado-prizma/json`)
            .then(r => r.json())
            .then(data => {{
                if (data.mensaje) msg.textContent = data.mensaje;
                if (data.total !== undefined) {{
                    const elTotal = document.getElementById("kpi-total");
                    const elExitosos = document.getElementById("kpi-exitosos");
                    const elPendientes = document.getElementById("kpi-pendientes");
                    const elFallidos = document.getElementById("kpi-fallidos");
                    if (elTotal) elTotal.textContent = data.total;
                    if (elExitosos) elExitosos.textContent = data.exitosos || 0;
                    if (elPendientes) elPendientes.textContent = data.pendientes !== undefined ? data.pendientes : (data.total - (data.procesados || 0));
                    if (elFallidos) elFallidos.textContent = data.fallidos || 0;

                    const proc = data.procesados || 0;
                    const tot = data.total || 0;
                    const pct = data.progreso_porcentaje !== undefined ? data.progreso_porcentaje : (tot > 0 ? Math.round((proc / tot) * 100) : 0);
                    const elPct = document.getElementById("progreso-porcentaje-texto");
                    const elBar = document.getElementById("progreso-barra-fill");
                    const elTxt = document.getElementById("progreso-texto-avance");
                    if (elPct) elPct.textContent = pct + "%";
                    if (elBar) elBar.style.width = pct + "%";
                    if (elTxt) elTxt.textContent = `Avance general (${{proc}} de ${{tot}} tests)`;
                    if (resumen) resumen.textContent = `${{proc}} de ${{tot}} test(s) procesados`;
                }}
                if (data.detalle_tests) {{
                    actualizarDetalleTests(data.detalle_tests, data.simulacion);
                }}

                // Manejo de selección interactiva de asignatura
                const pnlAsig = document.getElementById("panel-seleccion-asignatura");
                if (pnlAsig) {{
                    if (data.esperando_asignatura && data.opciones_asignatura && data.opciones_asignatura.length > 0) {{
                        pnlAsig.style.display = "flex";
                        const txtBusq = document.getElementById("txt-asignatura-buscada");
                        if (txtBusq && data.asignatura_buscada) {{
                            txtBusq.textContent = `La asignatura "${{data.asignatura_buscada}}" no coincide exactamente en PRIZMA. Selecciona la opción correcta entre las ${{data.opciones_asignatura.length}} disponibles para continuar:`;
                        }}
                        const selAsig = document.getElementById("sel-asignatura-opciones");
                        if (selAsig && selAsig.options.length <= 1) {{
                            selAsig.innerHTML = `<option value="">-- Seleccionar asignatura de PRIZMA (${{data.opciones_asignatura.length}} opciones) --</option>` +
                                data.opciones_asignatura.map(op => `<option value="${{op.replace(/"/g, '&quot;')}}">${{op}}</option>`).join("");
                        }}
                    }} else {{
                        pnlAsig.style.display = "none";
                    }}
                }}

                const btnDet = document.getElementById("btn-detener-auto");
                if (data.estado === "detenido") {{
                    badge.className = "badge warn";
                    badge.textContent = "⏹ Detenido por usuario";
                    const bDet = document.getElementById("banner-detenido-superior");
                    const txtDet = document.getElementById("texto-detenido-superior");
                    if (bDet) bDet.style.display = "flex";
                    if (txtDet && data.mensaje) txtDet.textContent = data.mensaje;
                    if (btnDet) btnDet.style.display = "none";
                    terminado = true;
                }} else if (data.estado === "completado") {{
                    badge.className = "badge ok";
                    badge.textContent = data.simulacion ? "✓ Simulación exitosa (Validado sin guardar)" : "✓ Guardado completado en PRIZMA";
                    const bOk = document.getElementById("banner-exito-superior");
                    const txtOk = document.getElementById("texto-exito-superior");
                    if (bOk) bOk.style.display = "flex";
                    if (txtOk && data.mensaje) txtOk.textContent = data.mensaje;
                    if (btnDet) btnDet.style.display = "none";
                    terminado = true;
                }} else if (data.estado === "error") {{
                    badge.className = "badge error";
                    badge.textContent = "❌ Error en la ejecución";
                    const bErr = document.getElementById("banner-error-superior");
                    const txtErr = document.getElementById("texto-error-superior");
                    if (bErr) bErr.style.display = "flex";
                    if (txtErr) txtErr.textContent = data.error || data.mensaje || "Error durante la automatización.";
                    if (btnDet) btnDet.style.display = "none";
                    if (bErr) bErr.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    terminado = true;
                }}
            }})
            .catch(() => {{}});
    }}

    const btnConfAsig = document.getElementById("btn-confirmar-asignatura");
    if (btnConfAsig) {{
        btnConfAsig.addEventListener("click", () => {{
            const selAsig = document.getElementById("sel-asignatura-opciones");
            const msgConf = document.getElementById("msg-confirmacion-asignatura");
            if (!selAsig || !selAsig.value) {{
                alert("Por favor selecciona una asignatura de la lista antes de continuar.");
                return;
            }}
            const formData = new FormData();
            formData.append("asignatura", selAsig.value);
            btnConfAsig.disabled = true;
            btnConfAsig.textContent = "⏳ Enviando selección...";
            fetch(`/tests/${{jobId}}/seleccionar-asignatura`, {{
                method: "POST",
                body: formData,
            }})
            .then(r => r.json())
            .then(res => {{
                if (res.ok) {{
                    if (msgConf) {{
                        msgConf.style.display = "block";
                        msgConf.style.color = "var(--verde)";
                        msgConf.textContent = `✓ Selección enviada: "${{res.asignatura}}". El robot continuará procesando el test...`;
                    }}
                    btnConfAsig.textContent = "✓ Confirmado";
                    setTimeout(() => {{
                        const pnl = document.getElementById("panel-seleccion-asignatura");
                        if (pnl) pnl.style.display = "none";
                    }}, 2500);
                }} else {{
                    alert(res.error || "Error al enviar la selección.");
                    btnConfAsig.disabled = false;
                    btnConfAsig.textContent = "Confirmar Asignatura y Continuar ➔";
                }}
            }})
            .catch(err => {{
                alert("Error de conexión al enviar selección: " + err);
                btnConfAsig.disabled = false;
                btnConfAsig.textContent = "Confirmar Asignatura y Continuar ➔";
            }});
        }});
    }}

    if (!terminado) {{
        setInterval(pollEstado, 1500);
    }}
    </script>'''
    return HTMLResponse(_layout("Monitoreo PRIZMA", body))


@router.get("/tests/{job_id}/estado-prizma/json")
def status_prizma_json(request: Request, job_id: str):
    job = _job_for(request, job_id)
    if not job:
        return JSONResponse({"error": "Análisis no encontrado"}, status_code=404)
    with _AUTOMATION_LOCK:
        auto_state = _AUTOMATION_JOBS.get(job_id)
    if not auto_state:
        return JSONResponse({"error": "Automatización no iniciada"}, status_code=404)
    return JSONResponse(auto_state)


@router.post("/tests/{job_id}/cancelar-prizma")
def cancelar_prizma(request: Request, job_id: str):
    job = _job_for(request, job_id)
    if not job:
        return JSONResponse({"error": "Análisis no encontrado"}, status_code=404)
    with _AUTOMATION_LOCK:
        auto_state = _AUTOMATION_JOBS.get(job_id)
    if not auto_state:
        return JSONResponse({"error": "Automatización no iniciada"}, status_code=404)
    if auto_state.get("estado") == "en_ejecucion":
        auto_state["cancelado"] = True
        auto_state["mensaje"] = "Deteniendo automatización a solicitud del usuario..."
        return JSONResponse({"ok": True, "mensaje": "Solicitud de detención registrada"})
    return JSONResponse({"ok": False, "mensaje": "La automatización no está en ejecución"})


@router.post("/tests/{job_id}/seleccionar-asignatura")
def seleccionar_asignatura(request: Request, job_id: str, asignatura: str = Form(...)):
    job = _job_for(request, job_id)
    if not job:
        return JSONResponse({"error": "Análisis no encontrado"}, status_code=404)
    with _AUTOMATION_LOCK:
        auto_state = _AUTOMATION_JOBS.get(job_id)
        sync_obj = _AUTOMATION_SYNC.get(job_id)
    if not auto_state or not sync_obj or not auto_state.get("esperando_asignatura"):
        return JSONResponse({"error": "No hay una selección de asignatura pendiente"}, status_code=400)

    asignatura_limpia = asignatura.strip()
    if not asignatura_limpia:
        return JSONResponse({"error": "Debes seleccionar una asignatura"}, status_code=400)

    sync_obj["seleccion"]["asignatura"] = asignatura_limpia
    sync_obj["event"].set()
    return JSONResponse({"ok": True, "asignatura": asignatura_limpia})


async def _extraer_bytes_docx_request(
    request: Request,
    archivo_docx: UploadFile | None,
    google_doc_url: str | None,
    modo_entrada: str,
) -> tuple[bytes, str, str | None]:
    """Retorna (contenido_bytes, nombre_fuente, error_str)."""
    modo_entrada = (modo_entrada or "archivo").strip().lower()
    if modo_entrada == "archivo" and archivo_docx and archivo_docx.filename:
        if not archivo_docx.filename.lower().endswith(".docx"):
            return b"", "", "El archivo debe tener formato Microsoft Word (.docx)."
        c = await archivo_docx.read()
        if not c:
            return b"", "", "El documento seleccionado está vacío o dañado."
        return c, archivo_docx.filename, None
    elif modo_entrada == "link" or (google_doc_url and google_doc_url.strip()):
        url_doc = str(google_doc_url or "").strip()
        if not url_doc:
            return b"", "", "Ingresa el enlace del Google Doc."
        loader = getattr(request.app.state, "test_google_loader", None)
        if not loader:
            return b"", "", "La conexión con Google Drive no está inicializada en el servidor."
        try:
            c, name = loader(url_doc, _user(request))
            if not c:
                return b"", "", "El documento de Google Docs está vacío."
            return c, name, None
        except Exception as exc:
            return b"", "", f"No se pudo descargar el Google Doc: {exc}"
    elif archivo_docx and archivo_docx.filename:
        if not archivo_docx.filename.lower().endswith(".docx"):
            return b"", "", "El archivo debe tener formato Microsoft Word (.docx)."
        c = await archivo_docx.read()
        if not c:
            return b"", "", "El documento seleccionado está vacío o dañado."
        return c, archivo_docx.filename, None
    return b"", "", "Selecciona un archivo Word (.docx) o introduce un enlace de Google Docs."


def _obtener_estado_test_activo(user: str = "") -> dict | None:
    """Solo devuelve un job mientras esta realmente en curso. En cuanto
    termina (exito o error) deja de aparecer aqui - el resultado se
    consulta en Historial/Reportes, no hace falta "marcarlo como visto"."""
    with _AUTOMATION_LOCK:
        for jid, astate in reversed(list(_AUTOMATION_JOBS.items())):
            if user and astate.get("owner") and astate.get("owner") != user:
                continue
            if astate.get("estado") in ["en_ejecucion", "esperando_asignatura"]:
                res = dict(astate)
                res["job_id"] = jid
                return res
    return None


@router.get("/tests/estado-activo/json")
def estado_test_activo_json(request: Request):
    user = _user(request)
    info = _obtener_estado_test_activo(user)
    if not info:
        return JSONResponse({"activo": False})
    return JSONResponse({
        "activo": True,
        "job_id": info.get("job_id"),
        "estado": info.get("estado"),
        "mensaje": info.get("mensaje"),
        "progreso_porcentaje": info.get("progreso_porcentaje", 0),
        "total": info.get("total", 0),
        "procesados": info.get("procesados", 0),
        "exitosos": info.get("exitosos", 0),
        "fallidos": info.get("fallidos", 0),
        "asignatura": info.get("asignatura", ""),
        "titulo": info.get("titulo_doc", ""),
    })
