from fastapi import (
    FastAPI,
    UploadFile,
    File,
    Form,
    BackgroundTasks,
)

from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    FileResponse,
)

from io import BytesIO

from openpyxl import load_workbook, Workbook

from motor_prizma import (
    ejecutar_cargue,
    validar_credenciales_prizma,
    determinar_tipo_archivo,
    normalizar_categoria,
    normalizar_texto,
)

import zipfile
import csv
import os
import uuid
import html
import json
import re
import threading
from collections import deque
from datetime import datetime
from zoneinfo import ZoneInfo


app = FastAPI(
    title="Auto Prizma Pro"
)


# ============================================================
# CARPETAS
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

UPLOADS_DIR = os.path.join(
    BASE_DIR,
    "uploads",
)

TEMP_DIR = os.path.join(
    BASE_DIR,
    "temp",
)

DATA_DIR = os.environ.get(
    "AUTO_PRIZMA_DATA_DIR",
    os.path.join(BASE_DIR, "data"),
)

RESULTADOS_DIR = os.path.join(
    DATA_DIR,
    "reportes",
)

HISTORIAL_PATH = os.path.join(
    DATA_DIR,
    "historial.json",
)


for carpeta in [
    UPLOADS_DIR,
    TEMP_DIR,
    DATA_DIR,
    RESULTADOS_DIR,
]:

    os.makedirs(
        carpeta,
        exist_ok=True,
    )


# ============================================================
# TRABAJOS
# ============================================================

TRABAJOS = {}
HISTORIAL_LOCK = threading.Lock()
TRABAJOS_LOCK = threading.RLock()
COLA_CARGUES = deque()
COLA_CONDICION = threading.Condition(TRABAJOS_LOCK)
ZONA_HORARIA_COLOMBIA = ZoneInfo("America/Bogota")


def _ahora_colombia():
    return datetime.now(ZONA_HORARIA_COLOMBIA)


def _ahora_colombia_iso():
    return _ahora_colombia().isoformat(timespec="seconds")


def _sanitizar_parte_nombre(valor, limite=70):
    texto = str(valor or "").strip()
    texto = re.sub(r'[\\/:*?"<>|]+', "-", texto)
    texto = re.sub(r"\s+", " ", texto).strip(" .-")
    if not texto:
        texto = "Sin nombre"
    return texto[:limite].rstrip()


def _cursos_desde_hojas(hojas):
    cursos = []
    vistos = set()

    for hoja in hojas or []:
        curso = str(hoja.get("curso") or "").strip() or "Curso sin nombre"
        programa = str(hoja.get("programa") or "").strip() or "Programa sin nombre"
        clave = (curso.casefold(), programa.casefold())
        if clave in vistos:
            continue
        vistos.add(clave)
        cursos.append({"curso": curso, "programa": programa})

    return cursos


def _nombre_reporte(cursos, trabajo_id):
    fecha = _ahora_colombia().strftime("%Y-%m-%d_%H-%M-%S")

    if cursos:
        curso = _sanitizar_parte_nombre(cursos[0]["curso"], 58)
        programa = _sanitizar_parte_nombre(cursos[0]["programa"], 58)
        extra = ""
        if len(cursos) > 1:
            extra = f"_y_{len(cursos) - 1}_curso_mas"
        return f"{fecha} - {curso} - {programa}{extra} - {trabajo_id[:6]}.csv"

    return f"{fecha} - resultado_prizma - {trabajo_id[:6]}.csv"


def _cargar_historial():
    if not os.path.isfile(HISTORIAL_PATH):
        return []

    try:
        with open(HISTORIAL_PATH, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
        return datos if isinstance(datos, list) else []
    except Exception:
        return []


def _guardar_historial(registros):
    temporal = HISTORIAL_PATH + ".tmp"
    with open(temporal, "w", encoding="utf-8") as archivo:
        json.dump(registros, archivo, ensure_ascii=False, indent=2)
    os.replace(temporal, HISTORIAL_PATH)


def _registrar_reporte_final(trabajo):
    ruta_reporte = trabajo.get("ruta_reporte")
    if not ruta_reporte or not os.path.isfile(ruta_reporte):
        return

    with HISTORIAL_LOCK:
        registros = _cargar_historial()
        trabajo_id = trabajo.get("id")

        if any(r.get("id") == trabajo_id for r in registros):
            trabajo["historial_registrado"] = True
            return

        fecha_iso = _ahora_colombia_iso()
        registros.append({
            "id": trabajo_id,
            "fecha_iso": fecha_iso,
            "archivo_reporte": os.path.basename(ruta_reporte),
            "cursos": trabajo.get("cursos", []),
            "estado_final": trabajo.get("etapa", "finalizado"),
        })
        _guardar_historial(registros)
        trabajo["historial_registrado"] = True


def ejecutar_cargue_con_historial(
    ruta_excel,
    ruta_zip,
    carpeta_temp,
    ruta_reporte,
    procesar_ovi,
    procesar_ova,
    procesar_retos,
    usuario_prizma,
    contrasena_prizma,
    trabajo,
):
    try:
        ejecutar_cargue(
            ruta_excel,
            ruta_zip,
            carpeta_temp,
            ruta_reporte,
            procesar_ovi,
            procesar_ova,
            procesar_retos,
            usuario_prizma,
            contrasena_prizma,
            trabajo,
        )
    finally:
        _registrar_reporte_final(trabajo)


# ============================================================
# COLA GLOBAL DE CARGUES - 1 PROCESO A LA VEZ
# ============================================================

def _posicion_en_cola(trabajo_id):
    with TRABAJOS_LOCK:
        pendientes = [
            tid for tid in COLA_CARGUES
            if tid in TRABAJOS
            and TRABAJOS[tid].get("etapa") == "en_cola"
            and not TRABAJOS[tid].get("terminado")
        ]
        try:
            return pendientes.index(trabajo_id) + 1
        except ValueError:
            return None


def _encolar_trabajo(trabajo_id, usuario_prizma, contrasena_prizma):
    with COLA_CONDICION:
        trabajo = TRABAJOS[trabajo_id]
        trabajo["usuario_prizma_temporal"] = usuario_prizma
        trabajo["contrasena_prizma_temporal"] = contrasena_prizma
        trabajo["etapa"] = "en_cola"
        trabajo["mensaje"] = "Cargue agregado a la cola de procesamiento."
        trabajo["encolado_en"] = _ahora_colombia_iso()
        trabajo["cancelar_solicitado"] = False
        trabajo["terminado"] = False
        if trabajo_id not in COLA_CARGUES:
            COLA_CARGUES.append(trabajo_id)
        COLA_CONDICION.notify_all()


def _worker_cola_cargues():
    while True:
        with COLA_CONDICION:
            while not COLA_CARGUES:
                COLA_CONDICION.wait()

            trabajo_id = COLA_CARGUES.popleft()
            trabajo = TRABAJOS.get(trabajo_id)
            if not trabajo:
                continue
            if trabajo.get("terminado") or trabajo.get("cancelar_solicitado"):
                continue

            usuario = trabajo.pop("usuario_prizma_temporal", "")
            contrasena = trabajo.pop("contrasena_prizma_temporal", "")
            trabajo["etapa"] = "iniciando"
            trabajo["mensaje"] = "Iniciando navegador y autenticación PRIZMA..."
            trabajo["iniciado_en"] = _ahora_colombia_iso()
            trabajo["terminado"] = False

        try:
            ejecutar_cargue_con_historial(
                trabajo["ruta_excel"],
                trabajo["ruta_zip"],
                trabajo["carpeta_temp"],
                trabajo["ruta_reporte"],
                trabajo["procesar_ovi"],
                trabajo["procesar_ova"],
                trabajo["procesar_retos"],
                usuario,
                contrasena,
                trabajo,
            )
        except Exception as exc:
            trabajo["etapa"] = "error"
            trabajo["mensaje"] = str(exc)
            trabajo["terminado"] = True
        finally:
            usuario = ""
            contrasena = ""
            trabajo.pop("usuario_prizma_temporal", None)
            trabajo.pop("contrasena_prizma_temporal", None)
            trabajo["finalizado_en"] = _ahora_colombia_iso()
            with COLA_CONDICION:
                COLA_CONDICION.notify_all()


_HILO_COLA = threading.Thread(
    target=_worker_cola_cargues,
    name="auto-prizma-cola",
    daemon=True,
)
_HILO_COLA.start()


# ============================================================
# CONVERTIR CSV A XLSX INTERNO
# ============================================================

def convertir_csv_a_xlsx(
    contenido_csv,
):

    texto = None

    for codificacion in [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ]:

        try:

            texto = contenido_csv.decode(
                codificacion
            )

            break

        except UnicodeDecodeError:
            continue

    if texto is None:
        raise ValueError(
            "No fue posible leer el archivo CSV."
        )

    muestra = texto[:10000]

    try:

        dialecto = csv.Sniffer().sniff(
            muestra,
            delimiters=",;\t|",
        )

        delimitador = dialecto.delimiter

    except csv.Error:

        delimitador = ","

    filas = list(
        csv.reader(
            texto.splitlines(),
            delimiter=delimitador,
        )
    )

    if not filas:
        raise ValueError(
            "El archivo CSV está vacío."
        )

    libro = Workbook()

    hoja = libro.active
    hoja.title = "Matriz"

    for numero_fila, fila in enumerate(
        filas,
        start=1,
    ):

        for numero_columna, valor in enumerate(
            fila,
            start=1,
        ):

            hoja.cell(
                row=numero_fila,
                column=numero_columna,
                value=valor,
            )

    salida = BytesIO()
    libro.save(salida)

    return salida.getvalue()


# ============================================================
# ANALIZAR MATRIZ
# ============================================================

def analizar_excel(
    contenido_excel,
    procesar_ovi=True,
    procesar_ova=True,
    procesar_retos=True,
):

    libro = load_workbook(
        BytesIO(contenido_excel),
        data_only=True,
    )

    hojas_validas = []

    for nombre_hoja in libro.sheetnames:

        hoja = libro[
            nombre_hoja
        ]

        fila_cabecera = None

        for fila in range(
            1,
            min(
                hoja.max_row,
                30,
            ) + 1,
        ):

            valor = hoja.cell(
                row=fila,
                column=1,
            ).value

            if normalizar_texto(
                valor
            ) == normalizar_texto(
                "Semana correspondiente"
            ):

                fila_cabecera = fila
                break

        if fila_cabecera is None:
            continue

        programa = hoja.cell(
            row=1,
            column=1,
        ).value

        curso = hoja.cell(
            row=2,
            column=1,
        ).value

        actividades = []

        cantidad_ovi = 0
        cantidad_ova = 0
        cantidad_retos = 0
        cantidad_h5p = 0
        cantidad_pdf = 0

        for fila in range(
            fila_cabecera + 1,
            hoja.max_row + 1,
        ):

            semana = hoja.cell(
                row=fila,
                column=1,
            ).value

            unidad = hoja.cell(
                row=fila,
                column=2,
            ).value

            nombre = hoja.cell(
                row=fila,
                column=4,
            ).value

            categoria = hoja.cell(
                row=fila,
                column=5,
            ).value

            tipo_recurso = hoja.cell(
                row=fila,
                column=6,
            ).value

            enlace = hoja.cell(
                row=fila,
                column=7,
            ).value

            if not nombre:
                continue

            if not categoria:
                continue

            categoria_prizma = normalizar_categoria(
                categoria
            )

            if categoria_prizma not in [
                "OVI",
                "OVA",
                "CHALLENGE",
            ]:
                continue

            if (
                categoria_prizma == "OVI"
                and not procesar_ovi
            ):
                continue

            if (
                categoria_prizma == "OVA"
                and not procesar_ova
            ):
                continue

            if (
                categoria_prizma == "CHALLENGE"
                and not procesar_retos
            ):
                continue

            tipo_archivo = determinar_tipo_archivo(
                categoria,
                tipo_recurso,
                enlace,
            )

            if tipo_archivo is None:
                continue

            if categoria_prizma == "OVI":
                cantidad_ovi += 1

            if categoria_prizma == "OVA":
                cantidad_ova += 1

            if categoria_prizma == "CHALLENGE":
                cantidad_retos += 1

            if tipo_archivo == "H5P":
                cantidad_h5p += 1

            if tipo_archivo == "PDF":
                cantidad_pdf += 1

            actividades.append(
                {
                    "fila": fila,

                    "semana": (
                        str(semana).strip()
                        if semana
                        else ""
                    ),

                    "unidad": (
                        str(unidad).strip()
                        if unidad
                        else ""
                    ),

                    "nombre": str(
                        nombre
                    ).strip(),

                    "categoria":
                        categoria_prizma,

                    "tipo":
                        tipo_archivo,
                }
            )

        hojas_validas.append(
            {
                "hoja":
                    nombre_hoja,

                "programa": (
                    str(programa).strip()
                    if programa
                    else ""
                ),

                "curso": (
                    str(curso).strip()
                    if curso
                    else nombre_hoja
                ),

                "actividades":
                    actividades,

                "ovi":
                    cantidad_ovi,

                "ova":
                    cantidad_ova,

                "retos":
                    cantidad_retos,

                "h5p":
                    cantidad_h5p,

                "pdf":
                    cantidad_pdf,
            }
        )

    return hojas_validas


# ============================================================
# ANALIZAR ZIP
# ============================================================

def analizar_zip(
    contenido_zip,
):

    cantidad_h5p = 0
    cantidad_pdf = 0

    with zipfile.ZipFile(
        BytesIO(contenido_zip),
        "r",
    ) as archivo_zip:

        for miembro in archivo_zip.namelist():

            if miembro.endswith("/"):
                continue

            nombre = os.path.basename(
                miembro
            )

            extension = os.path.splitext(
                nombre
            )[1].lower()

            if extension == ".h5p":

                cantidad_h5p += 1

            elif extension == ".pdf":

                cantidad_pdf += 1

    return {
        "h5p":
            cantidad_h5p,

        "pdf":
            cantidad_pdf,

        "total":
            cantidad_h5p
            + cantidad_pdf,
    }


# ============================================================
# HTML PRINCIPAL
# ============================================================

def generar_html(
    resultado=None,
    error=None,
    trabajo_id=None,
):

    def e(valor):
        return html.escape(
            str(valor or "")
        )

    contenido = ""

    # ========================================================
    # PANTALLA INICIAL
    # ========================================================

    if resultado is None:

        bloque_error = ""

        if error:
            bloque_error = f"""
            <div class="alerta-error">
                <div class="alerta-icono">!</div>
                <div>
                    <strong>No se pudo continuar</strong>
                    <p>{e(error)}</p>
                </div>
            </div>
            """

        contenido = f"""
        <section class="encabezado-pagina">
            <div>
                <h1>¡Bienvenido!</h1>
                <p>Automatiza el cargue de actividades en PRIZMA de forma rápida y segura.</p>
            </div>
        </section>

        {bloque_error}

        <form
            id="cargue-principal"
            action="/analizar"
            method="post"
            enctype="multipart/form-data"
            class="panel panel-principal"
        >
            <div class="titulo-seccion">
                <span class="numero-seccion">1</span>
                <div>
                    <h2>Archivos</h2>
                    <p>Selecciona la matriz y el paquete de recursos.</p>
                </div>
            </div>

            <div class="grid-archivos">
                <label class="tarjeta-archivo zona-drop" for="archivo-matriz" data-input="archivo-matriz">
                    <div class="archivo-cabecera">
                        <div class="archivo-icono verde" aria-hidden="true">
                            <svg viewBox="0 0 24 24"><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5"/><path d="M9 11l6 6M15 11l-6 6"/></svg>
                        </div>
                        <div>
                            <strong>Archivo Excel o CSV</strong>
                            <span>Formatos .xlsx y .csv · también puedes arrastrarlo aquí</span>
                        </div>
                    </div>
                    <div class="selector-archivo">
                        <span class="boton-selector">Seleccionar archivo</span>
                        <span id="nombre-matriz" class="nombre-archivo">Ningún archivo seleccionado</span>
                    </div>
                    <input
                        id="archivo-matriz"
                        type="file"
                        name="excel"
                        accept=".xlsx,.csv"
                        required
                    >
                </label>

                <label class="tarjeta-archivo zona-drop" for="archivo-zip" data-input="archivo-zip">
                    <div class="archivo-cabecera">
                        <div class="archivo-icono morado" aria-hidden="true">
                            <svg viewBox="0 0 24 24"><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5"/><path d="M11 5h2M11 8h2M11 11h2M11 14h2"/></svg>
                        </div>
                        <div>
                            <strong>ZIP de recursos</strong>
                            <span>Archivos H5P y PDF · también puedes arrastrarlo aquí</span>
                        </div>
                    </div>
                    <div class="selector-archivo">
                        <span class="boton-selector">Seleccionar archivo</span>
                        <span id="nombre-zip" class="nombre-archivo">Ningún archivo seleccionado</span>
                    </div>
                    <input
                        id="archivo-zip"
                        type="file"
                        name="recursos"
                        accept=".zip"
                        required
                    >
                </label>
            </div>

            <div class="separador"></div>

            <div class="titulo-seccion">
                <span class="numero-seccion">2</span>
                <div>
                    <h2>Tipos de actividades</h2>
                    <p>Selecciona qué actividades deseas procesar.</p>
                </div>
            </div>

            <div class="grid-tipos">
                <label class="tarjeta-tipo tipo-verde">
                    <input type="checkbox" name="ovi" value="1" checked>
                    <span class="check-personalizado">✓</span>
                    <span class="tipo-icono icono-ovi" aria-hidden="true">
                        <svg viewBox="0 0 24 24"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H11v16H6.5A2.5 2.5 0 0 0 4 21.5z"/><path d="M20 5.5A2.5 2.5 0 0 0 17.5 3H13v16h4.5a2.5 2.5 0 0 1 2.5 2.5z"/></svg>
                    </span>
                    <strong>OVI</strong>
                    <small>Objetos Virtuales de Información</small>
                </label>

                <label class="tarjeta-tipo tipo-azul">
                    <input type="checkbox" name="ova" value="1" checked>
                    <span class="check-personalizado">✓</span>
                    <span class="tipo-icono icono-ova" aria-hidden="true">
                        <svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4"/><path d="m9 9 2 2 4-4"/></svg>
                    </span>
                    <strong>OVA</strong>
                    <small>Objetos Virtuales de Aprendizaje</small>
                </label>

                <label class="tarjeta-tipo tipo-morado">
                    <input type="checkbox" name="retos" value="1" checked>
                    <span class="check-personalizado">✓</span>
                    <span class="tipo-icono icono-reto" aria-hidden="true">
                        <svg viewBox="0 0 24 24"><path d="M8 3h8v4a4 4 0 0 1-8 0z"/><path d="M8 5H4v2a4 4 0 0 0 4 4M16 5h4v2a4 4 0 0 1-4 4"/><path d="M12 11v5M8 21h8M10 16h4v5h-4z"/></svg>
                    </span>
                    <strong>Retos Evaluativos</strong>
                    <small>Retos evaluativos en PDF</small>
                </label>
            </div>

            <button class="boton-principal" type="submit">
                <span>↗</span>
                Analizar archivos
            </button>

            <p class="nota-seguridad">
                Tus archivos se utilizan únicamente durante el proceso de cargue.
            </p>
        </form>

        <aside id="ayuda" class="panel panel-ayuda">
            <h3>¿Cómo funciona?</h3>
            <div class="paso"><b>1</b><span>Sube tu Excel o CSV con las actividades.</span></div>
            <div class="paso"><b>2</b><span>Sube el ZIP con los recursos correspondientes.</span></div>
            <div class="paso"><b>3</b><span>Selecciona los tipos de actividades.</span></div>
            <div class="paso"><b>4</b><span>Analiza, revisa y luego inicia el cargue.</span></div>

            <div class="ayuda-separador"></div>

            <h3 class="titulo-consejos">Consejos</h3>
            <p class="consejo">✓ Verifica que la matriz conserve el formato esperado.</p>
            <p class="consejo">✓ El ZIP debe contener los H5P y PDF correspondientes.</p>
            <p class="consejo">✓ Revisa el análisis antes de iniciar el cargue.</p>
        </aside>
        """

    # ========================================================
    # PANTALLA DE REVISIÓN
    # ========================================================

    else:

        total_actividades = 0
        total_ovi = 0
        total_ova = 0
        total_retos = 0
        total_h5p = 0
        total_pdf = 0
        filas_html = ""

        for hoja in resultado["hojas"]:

            total_actividades += len(
                hoja["actividades"]
            )
            total_ovi += hoja.get("ovi", 0)
            total_ova += hoja.get("ova", 0)
            total_retos += hoja.get("retos", 0)
            total_h5p += hoja.get("h5p", 0)
            total_pdf += hoja.get("pdf", 0)

            for actividad in hoja["actividades"]:

                categoria = actividad["categoria"]
                clase_categoria = {
                    "OVI": "badge-verde",
                    "OVA": "badge-azul",
                    "CHALLENGE": "badge-morado",
                }.get(categoria, "badge-gris")

                etiqueta_categoria = (
                    "RETO"
                    if categoria == "CHALLENGE"
                    else categoria
                )

                clase_tipo = (
                    "badge-pdf"
                    if actividad["tipo"] == "PDF"
                    else "badge-h5p"
                )

                filas_html += f"""
                <tr>
                    <td>{e(actividad["semana"])}</td>
                    <td>{e(actividad["unidad"])}</td>
                    <td class="celda-actividad">{e(actividad["nombre"])}</td>
                    <td><span class="badge {clase_categoria}">{e(etiqueta_categoria)}</span></td>
                    <td><span class="badge {clase_tipo}">{e(actividad["tipo"])}</span></td>
                </tr>
                """

        bloque_error_revision = ""
        if error:
            bloque_error_revision = f"""
            <div class="alerta-error" style="grid-column:1/-1;margin-bottom:18px;">
                <div class="alerta-icono">!</div>
                <div><strong>No se pudo iniciar el cargue</strong><p>{e(error)}</p></div>
            </div>
            """

        contenido = f"""
        <section id="cargue-principal" class="encabezado-exito panel">
            <div class="check-grande">✓</div>
            <div>
                <h1>Análisis completado</h1>
                <p>{total_actividades} actividades listas para cargar</p>
            </div>
        </section>

        <section class="resumen-grid">
            <div class="resumen-card"><span>Actividades</span><strong>{total_actividades}</strong></div>
            <div class="resumen-card verde"><span>OVI</span><strong>{total_ovi}</strong></div>
            <div class="resumen-card azul"><span>OVA</span><strong>{total_ova}</strong></div>
            <div class="resumen-card morado"><span>Retos</span><strong>{total_retos}</strong></div>
            <div class="resumen-card violeta"><span>H5P</span><strong>{total_h5p}</strong></div>
            <div class="resumen-card rojo"><span>PDF</span><strong>{total_pdf}</strong></div>
        </section>

        {bloque_error_revision}

        <section class="revision-grid">
            <div class="panel tabla-panel">
                <div class="titulo-bloque">
                    <h2>Actividades detectadas</h2>
                    <p>Revisa que las actividades y categorías sean correctas.</p>
                </div>

                <div class="tabla-scroll">
                    <table>
                        <thead>
                            <tr>
                                <th>Semana</th>
                                <th>Unidad</th>
                                <th>Actividad</th>
                                <th>Categoría</th>
                                <th>Tipo</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filas_html}
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="panel credenciales-panel">
                <div class="titulo-bloque">
                    <h2>Acceso a PRIZMA</h2>
                    <p>Ingresa tus credenciales para autorizar este cargue.</p>
                </div>

                <form action="/iniciar/{trabajo_id}" method="post" autocomplete="off">
                    <label class="campo-label" for="usuario_prizma">Usuario PRIZMA</label>
                    <div class="campo-moderno">
                        <span>♙</span>
                        <input
                            id="usuario_prizma"
                            type="text"
                            name="usuario_prizma"
                            placeholder="Ingresa tu usuario"
                            autocomplete="off"
                            required
                        >
                    </div>

                    <label class="campo-label" for="contrasena_prizma">Contraseña PRIZMA</label>
                    <div class="campo-moderno">
                        <span>⌾</span>
                        <input
                            id="contrasena_prizma"
                            type="password"
                            name="contrasena_prizma"
                            placeholder="Ingresa tu contraseña"
                            autocomplete="new-password"
                            required
                        >
                    </div>

                    <div class="caja-seguridad">
                        <strong>Credenciales de uso temporal</strong>
                        <span>La aplicación no las escribe en archivos, reportes ni base de datos.</span>
                    </div>

                    <button class="boton-principal" type="submit">
                        <span>↗</span>
                        Iniciar cargue en PRIZMA
                    </button>
                </form>
            </div>
        </section>

        <a class="boton-secundario" href="/">← Cambiar archivos</a>
        """

    return """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Auto Prizma Pro</title>
        <style>
            :root {
                --fondo: #f7f8fc;
                --panel: #ffffff;
                --texto: #101828;
                --muted: #667085;
                --borde: #e5e7ef;
                --morado: #5b48e8;
                --morado-2: #4338ca;
                --verde: #0ea968;
                --azul: #1976d2;
                --rojo: #ef4444;
                --sombra: 0 10px 30px rgba(29, 41, 57, .06);
            }

            * { box-sizing: border-box; }

            body {
                margin: 0;
                font-family: Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                background: var(--fondo);
                color: var(--texto);
            }

            .app { min-height: 100vh; display: grid; grid-template-columns: 235px 1fr; }

            .sidebar {
                position: sticky;
                top: 0;
                height: 100vh;
                background: #fff;
                border-right: 1px solid var(--borde);
                padding: 28px 20px;
                display: flex;
                flex-direction: column;
            }

            .marca { display: flex; align-items: center; gap: 12px; margin-bottom: 34px; }
            .logo {
                width: 44px; height: 44px; border-radius: 13px;
                display: grid; place-items: center;
                background: linear-gradient(145deg, #6d5dfc, #4338ca);
                box-shadow: 0 8px 20px rgba(79, 70, 229, .25);
                flex: 0 0 44px;
            }
            .logo svg { width: 29px; height: 29px; overflow: visible; }
            .logo svg path:first-child { fill: #fff; }
            .logo svg path:last-child { fill: #c7d2fe; }
            .marca strong { display: block; font-size: 18px; }
            .marca span { display: block; color: var(--muted); font-size: 12px; margin-top: 3px; }

            .nav { display: grid; gap: 8px; }
            .nav-item {
                padding: 12px 14px; border-radius: 11px; color: #475467; font-size: 14px;
                display: flex; gap: 11px; align-items: center; text-decoration: none; transition: .18s ease;
            }
            a.nav-item:hover { background: #f7f5ff; color: #4f46e5; }
            .nav-item.activo { background: #f1efff; color: #4f46e5; font-weight: 700; }
            .nav-item.proximamente { opacity: .48; cursor: default; }
            .nav-item.proximamente small { margin-left: auto; font-size: 9px; }

            .estado-servicio {
                margin-top: auto; border: 1px solid var(--borde); border-radius: 14px;
                padding: 15px; background: #fff;
            }
            .servicio-linea { display: flex; align-items: center; gap: 8px; color: #07894f; font-size: 13px; font-weight: 700; }
            .punto { width: 8px; height: 8px; border-radius: 50%; background: #15b76a; }
            .servicio-mini { margin-top: 13px; display: flex; justify-content: space-between; font-size: 12px; color: var(--muted); }
            .chip { background: #eef4ff; color: #3538cd; padding: 4px 8px; border-radius: 999px; }

            .contenido { padding: 28px 34px 45px; max-width: 1500px; width: 100%; margin: 0 auto; }
            .layout-inicio { display: grid; grid-template-columns: minmax(0, 1fr) 290px; gap: 24px; align-items: start; }

            .panel {
                background: var(--panel); border: 1px solid var(--borde); border-radius: 16px;
                box-shadow: var(--sombra);
            }

            .encabezado-pagina {
                grid-column: 1 / -1; background: #fff; border: 1px solid var(--borde);
                border-radius: 16px; padding: 24px 28px; box-shadow: var(--sombra);
            }
            .encabezado-pagina h1, .encabezado-exito h1 { margin: 0; font-size: 25px; }
            .encabezado-pagina p, .encabezado-exito p { margin: 6px 0 0; color: var(--muted); }

            .panel-principal { padding: 26px; }
            .titulo-seccion { display: flex; gap: 12px; align-items: flex-start; margin-bottom: 18px; }
            .titulo-seccion h2, .titulo-bloque h2 { margin: 0; font-size: 18px; }
            .titulo-seccion p, .titulo-bloque p { margin: 5px 0 0; color: var(--muted); font-size: 13px; }
            .numero-seccion {
                width: 28px; height: 28px; border-radius: 9px; display: grid; place-items: center;
                background: #f0edff; color: #5145cd; font-weight: 800; font-size: 13px;
            }

            .grid-archivos { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
            .tarjeta-archivo {
                border: 1px dashed #cfd4e1; border-radius: 14px; padding: 20px;
                cursor: pointer; transition: .18s ease; background: #fff;
                min-width: 0; overflow: hidden;
            }
            .tarjeta-archivo:hover { border-color: #8176f2; background: #fbfaff; transform: translateY(-1px); }
            .tarjeta-archivo.arrastrando { border-color: #5b48e8; background: #f5f3ff; box-shadow: inset 0 0 0 2px rgba(91,72,232,.10); }
            .tarjeta-archivo input { display: none; }
            .archivo-cabecera { display: flex; align-items: center; gap: 12px; min-width: 0; }
            .archivo-cabecera > div:last-child { min-width: 0; }
            .archivo-cabecera strong { display: block; font-size: 14px; }
            .archivo-cabecera span { display: block; font-size: 12px; color: var(--muted); margin-top: 4px; }
            .archivo-icono {
                width: 40px; height: 40px; border-radius: 10px; display: grid; place-items: center;
                flex: 0 0 40px;
            }
            .archivo-icono svg { width: 23px; height: 23px; fill: none; stroke: currentColor; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }
            .archivo-icono.verde { background: #e8f8f0; color: #0a9b5b; }
            .archivo-icono.morado { background: #f0edff; color: #5b48e8; }
            .selector-archivo {
                margin-top: 18px; min-height: 46px; border: 1px solid var(--borde); border-radius: 10px;
                padding: 9px 12px; display: flex; align-items: center; gap: 12px;
                min-width: 0; overflow: hidden;
            }
            .boton-selector {
                background: #5b48e8; color: #fff; padding: 8px 12px; border-radius: 8px;
                font-size: 12px; font-weight: 700; white-space: nowrap;
            }
            .nombre-archivo { color: var(--muted); font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; flex: 1 1 auto; display: block; }

            .separador { height: 1px; background: var(--borde); margin: 26px 0; }
            .grid-tipos { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
            .tarjeta-tipo {
                position: relative; border: 1px solid var(--borde); border-radius: 14px; padding: 18px;
                min-height: 125px; cursor: pointer; transition: .18s ease; display: flex; flex-direction: column;
            }
            .tarjeta-tipo:hover { transform: translateY(-1px); box-shadow: 0 8px 22px rgba(16,24,40,.06); }
            .tarjeta-tipo input { position: absolute; opacity: 0; pointer-events: none; }
            .check-personalizado {
                position: absolute; top: 15px; right: 15px; width: 22px; height: 22px;
                border-radius: 6px; display: grid; place-items: center; color: white; font-size: 13px;
                background: #d0d5dd;
            }
            .tarjeta-tipo:has(input:checked).tipo-verde { border-color: #65d6a1; background: #fbfffd; }
            .tarjeta-tipo:has(input:checked).tipo-verde .check-personalizado { background: #0ea968; }
            .tarjeta-tipo:has(input:checked).tipo-azul { border-color: #7ab6ef; background: #fbfdff; }
            .tarjeta-tipo:has(input:checked).tipo-azul .check-personalizado { background: #1976d2; }
            .tarjeta-tipo:has(input:checked).tipo-morado { border-color: #b9a8f5; background: #fdfcff; }
            .tarjeta-tipo:has(input:checked).tipo-morado .check-personalizado { background: #8b5cf6; }
            .tipo-icono { width: 42px; height: 42px; border-radius: 12px; display: grid; place-items: center; margin-bottom: 15px; }
            .tipo-icono svg { width: 25px; height: 25px; fill: none; stroke: currentColor; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }
            .icono-ovi { color:#079455; background:#e9f9f1; }
            .icono-ova { color:#1976d2; background:#eaf4ff; }
            .icono-reto { color:#7f56d9; background:#f1ebff; }
            .tarjeta-tipo strong { font-size: 16px; }
            .tarjeta-tipo small { color: var(--muted); margin-top: 6px; line-height: 1.4; }

            .boton-principal {
                width: 100%; border: 0; border-radius: 11px; padding: 15px 20px; margin-top: 24px;
                color: #fff; font-size: 15px; font-weight: 800; cursor: pointer;
                background: linear-gradient(90deg, #5145e5, #4f46e5 55%, #5f45df);
                box-shadow: 0 8px 20px rgba(79,70,229,.20);
            }
            .boton-principal:hover { filter: brightness(.98); transform: translateY(-1px); }
            .boton-principal span { margin-right: 8px; }
            .nota-seguridad { text-align: center; color: var(--muted); font-size: 12px; margin: 14px 0 0; }

            .panel-ayuda { padding: 24px; }
            .panel-ayuda h3 { margin: 0 0 20px; font-size: 17px; }
            .paso { display: grid; grid-template-columns: 30px 1fr; gap: 12px; align-items: start; margin-bottom: 20px; color: #475467; font-size: 13px; line-height: 1.5; }
            .paso b { width: 28px; height: 28px; border-radius: 50%; background: #f1efff; color: #4f46e5; display: grid; place-items: center; }
            .ayuda-separador { height: 1px; background: var(--borde); margin: 24px 0; }
            .titulo-consejos { color: #4f46e5; }
            .consejo { color: #475467; font-size: 13px; line-height: 1.55; }

            .alerta-error {
                grid-column: 1 / -1; display: flex; gap: 12px; padding: 16px 18px; border-radius: 12px;
                background: #fff1f2; border: 1px solid #fecdd3; color: #9f1239;
            }
            .alerta-error p { margin: 4px 0 0; }
            .alerta-icono { width: 28px; height: 28px; border-radius: 50%; background: #ef4444; color: #fff; display: grid; place-items: center; font-weight: 900; }

            .encabezado-exito { padding: 22px 24px; display: flex; gap: 15px; align-items: center; margin-bottom: 0; grid-column: 1 / -1; }
            .check-grande { width: 48px; height: 48px; border-radius: 14px; background: #e9f9f1; color: #08a45e; display: grid; place-items: center; font-size: 25px; font-weight: 900; }
            .resumen-grid { display: grid; grid-template-columns: repeat(6, minmax(110px, 1fr)); gap: 14px; margin: 18px 0 20px; grid-column: 1 / -1; min-width: 0; }
            .resumen-card { background: #fff; border: 1px solid var(--borde); border-radius: 14px; padding: 17px; box-shadow: var(--sombra); text-align: center; }
            .resumen-card span { display: block; color: var(--muted); font-size: 12px; }
            .resumen-card strong { display: block; font-size: 27px; margin-top: 6px; }
            .resumen-card.verde strong { color: #0a9b5b; }
            .resumen-card.azul strong { color: #1976d2; }
            .resumen-card.morado strong, .resumen-card.violeta strong { color: #6d4ce8; }
            .resumen-card.rojo strong { color: #e5484d; }

            .revision-grid { display: grid; grid-template-columns: minmax(0, 1.55fr) minmax(310px, .75fr); gap: 20px; align-items: start; grid-column: 1 / -1; min-width: 0; }
            .tabla-panel, .credenciales-panel { padding: 24px; }
            .titulo-bloque { margin-bottom: 18px; }
            .tabla-scroll { max-height: 570px; overflow: auto; border: 1px solid var(--borde); border-radius: 12px; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            thead { position: sticky; top: 0; background: #f8f9fc; z-index: 1; }
            th { text-align: left; color: #475467; font-size: 12px; padding: 12px 13px; border-bottom: 1px solid var(--borde); }
            td { padding: 12px 13px; border-bottom: 1px solid #eef0f4; vertical-align: middle; }
            tbody tr:last-child td { border-bottom: 0; }
            .celda-actividad { min-width: 260px; }
            .badge { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 10px; font-weight: 800; }
            .badge-verde { background: #e9f9f1; color: #07894f; }
            .badge-azul { background: #eaf4ff; color: #1565c0; }
            .badge-morado { background: #f1ebff; color: #7048d7; }
            .badge-gris { background: #f2f4f7; color: #475467; }
            .badge-h5p { background: #f1ebff; color: #6941c6; }
            .badge-pdf { background: #fff0f1; color: #d92d20; }

            .campo-label { display: block; font-size: 13px; font-weight: 700; margin: 18px 0 8px; }
            .campo-moderno { display: flex; align-items: center; gap: 10px; border: 1px solid #d7dbe5; border-radius: 10px; padding: 0 12px; background: #fff; }
            .campo-moderno span { color: #667085; }
            .campo-moderno input { width: 100%; border: 0; outline: 0; padding: 13px 0; font-size: 14px; background: transparent; }
            .caja-seguridad { margin-top: 20px; background: #f8f7ff; border: 1px solid #d9d4ff; border-radius: 11px; padding: 14px; }
            .caja-seguridad strong, .caja-seguridad span { display: block; }
            .caja-seguridad strong { font-size: 12px; color: #4f46e5; }
            .caja-seguridad span { font-size: 12px; color: var(--muted); margin-top: 5px; line-height: 1.45; }
            .boton-secundario { grid-column: 1 / -1; display: inline-block; margin-top: 20px; color: #4f46e5; text-decoration: none; font-weight: 700; font-size: 14px; padding: 11px 15px; border: 1px solid #d8d6f8; border-radius: 10px; background: #fff; }

            @media (max-width: 1050px) {
                .app { grid-template-columns: 1fr; }
                .sidebar { display: none; }
                .contenido { padding: 20px; }
                .layout-inicio { grid-template-columns: 1fr; }
                .panel-ayuda { order: 2; }
                .resumen-grid { grid-template-columns: repeat(3, 1fr); }
                .revision-grid { grid-template-columns: 1fr; }
            }

            @media (max-width: 700px) {
                .grid-archivos, .grid-tipos { grid-template-columns: 1fr; }
                .resumen-grid { grid-template-columns: repeat(2, 1fr); }
                .contenido { padding: 14px; }
                .panel-principal, .tabla-panel, .credenciales-panel { padding: 18px; }
            }
        </style>
    </head>
    <body>
        <div class="app">
            <aside class="sidebar">
                <div class="marca">
                    <div class="logo" aria-label="Auto Prizma Pro">
                        <svg viewBox="0 0 48 48" aria-hidden="true">
                            <path d="M9 35.5 20.5 8.5c.8-1.9 3.4-1.9 4.2 0l4.1 9.6-5.4 12.7-3.1-7.4-5.2 12.1z"/>
                            <path d="M26.4 14.5 39 35.5h-8.2l-8.5-14.2z"/>
                        </svg>
                    </div>
                    <div>
                        <strong>Auto Prizma Pro</strong>
                        <span>Automatización PRIZMA</span>
                    </div>
                </div>

                <nav class="nav">
                    <a class="nav-item activo" href="/">⌂ <span>Inicio</span></a>
                    <a class="nav-item" href="/cargue-actual">⇧ <span>Cargue actual</span></a>
                    <a class="nav-item" href="/historial">◷ <span>Historial</span></a>
                    <a class="nav-item" href="/reportes">▥ <span>Reportes</span></a>
                </nav>

                <div class="estado-servicio">
                    <div class="servicio-linea"><span class="punto"></span> Servicio activo</div>
                    <div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div>
                    <div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div>
                </div>
            </aside>

            <main class="contenido">
                <div class="layout-inicio">
                    """ + contenido + """
                </div>
            </main>
        </div>

        <script>
            const matriz = document.getElementById('archivo-matriz');
            const zip = document.getElementById('archivo-zip');

            function actualizarNombre(input, destinoId) {
                const destino = document.getElementById(destinoId);
                if (!destino || !input) return;
                destino.textContent = input.files.length ? input.files[0].name : 'Ningún archivo seleccionado';
                destino.title = input.files.length ? input.files[0].name : '';
            }

            if (matriz) matriz.addEventListener('change', () => actualizarNombre(matriz, 'nombre-matriz'));
            if (zip) zip.addEventListener('change', () => actualizarNombre(zip, 'nombre-zip'));

            document.querySelectorAll('.zona-drop').forEach((zona) => {
                const input = document.getElementById(zona.dataset.input);
                if (!input) return;

                ['dragenter', 'dragover'].forEach((evento) => {
                    zona.addEventListener(evento, (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        zona.classList.add('arrastrando');
                    });
                });

                ['dragleave', 'drop'].forEach((evento) => {
                    zona.addEventListener(evento, (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        zona.classList.remove('arrastrando');
                    });
                });

                zona.addEventListener('drop', (e) => {
                    const archivos = e.dataTransfer.files;
                    if (!archivos || !archivos.length) return;
                    const transferencia = new DataTransfer();
                    transferencia.items.add(archivos[0]);
                    input.files = transferencia.files;
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                });
            });
        </script>
    </body>
    </html>
    """


# ============================================================
# INICIO
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse,
)
def inicio():

    return generar_html()


# ============================================================
# ANALIZAR
# ============================================================

@app.post(
    "/analizar",
    response_class=HTMLResponse,
)
async def analizar(
    excel: UploadFile = File(...),
    recursos: UploadFile = File(...),
    ovi: str | None = Form(default=None),
    ova: str | None = Form(default=None),
    retos: str | None = Form(default=None),
):

    try:

        nombre_matriz = (
            excel.filename
            or ""
        ).lower()

        if not nombre_matriz.endswith(
            (".xlsx", ".csv")
        ):

            return generar_html(
                error=(
                    "La matriz debe ser un archivo "
                    ".xlsx o .csv"
                )
            )

        if not recursos.filename.lower().endswith(
            ".zip"
        ):

            return generar_html(
                error=(
                    "Los recursos deben estar "
                    "en un archivo .zip"
                )
            )

        procesar_ovi = (
            ovi is not None
        )

        procesar_ova = (
            ova is not None
        )

        procesar_retos = (
            retos is not None
        )

        if (
            not procesar_ovi
            and not procesar_ova
            and not procesar_retos
        ):

            return generar_html(
                error=(
                    "Selecciona OVI, OVA y/o Retos Evaluativos."
                )
            )

        contenido_matriz = await excel.read()

        contenido_zip = await recursos.read()

        if nombre_matriz.endswith(
            ".csv"
        ):

            contenido_excel = convertir_csv_a_xlsx(
                contenido_matriz
            )

        else:

            contenido_excel = contenido_matriz

        hojas = analizar_excel(
            contenido_excel,
            procesar_ovi,
            procesar_ova,
            procesar_retos,
        )

        if not hojas:

            return generar_html(
                error=(
                    "No encontré una matriz "
                    "PRIZMA válida."
                )
            )

        try:

            zip_info = analizar_zip(
                contenido_zip
            )

        except zipfile.BadZipFile:

            return generar_html(
                error=(
                    "El ZIP no es válido."
                )
            )

        # ====================================================
        # CREAR TRABAJO
        # ====================================================

        trabajo_id = uuid.uuid4().hex

        carpeta_trabajo = os.path.join(
            UPLOADS_DIR,
            trabajo_id,
        )

        os.makedirs(
            carpeta_trabajo,
            exist_ok=True,
        )

        ruta_excel = os.path.join(
            carpeta_trabajo,
            "matriz.xlsx",
        )

        ruta_zip = os.path.join(
            carpeta_trabajo,
            "recursos.zip",
        )

        with open(
            ruta_excel,
            "wb",
        ) as archivo:

            archivo.write(
                contenido_excel
            )

        with open(
            ruta_zip,
            "wb",
        ) as archivo:

            archivo.write(
                contenido_zip
            )

        carpeta_temp = os.path.join(
            TEMP_DIR,
            trabajo_id,
        )

        os.makedirs(
            carpeta_temp,
            exist_ok=True,
        )

        cursos_trabajo = _cursos_desde_hojas(hojas)

        ruta_reporte = os.path.join(
            RESULTADOS_DIR,
            _nombre_reporte(cursos_trabajo, trabajo_id),
        )

        TRABAJOS[
            trabajo_id
        ] = {
            "id":
                trabajo_id,

            "ruta_excel":
                ruta_excel,

            "ruta_zip":
                ruta_zip,

            "carpeta_temp":
                carpeta_temp,

            "ruta_reporte":
                ruta_reporte,

            "procesar_ovi":
                procesar_ovi,

            "procesar_ova":
                procesar_ova,

            "procesar_retos":
                procesar_retos,

            "etapa":
                "analizado",

            "mensaje":
                "Archivos analizados correctamente.",

            "total":
                0,

            "procesadas":
                0,

            "exitosas":
                0,

            "errores":
                0,

            "terminado":
                False,

            "detalle_actividades":
                [],

            "cursos":
                cursos_trabajo,

            "historial_registrado":
                False,

            "creado_en":
                _ahora_colombia_iso(),

            "cancelar_solicitado":
                False,
        }

        resultado = {
            "hojas":
                hojas,

            "zip":
                zip_info,
        }

        TRABAJOS[trabajo_id]["resultado_analisis"] = resultado

        return generar_html(
            resultado=resultado,
            trabajo_id=trabajo_id,
        )

    except Exception as e:

        return generar_html(
            error=str(e)
        )


# ============================================================
# INICIAR
# ============================================================

@app.post(
    "/iniciar/{trabajo_id}",
    response_class=HTMLResponse,
)
def iniciar_trabajo(
    trabajo_id: str,
    background_tasks: BackgroundTasks,
    usuario_prizma: str = Form(...),
    contrasena_prizma: str = Form(...),
):
    trabajo = TRABAJOS.get(trabajo_id)
    if not trabajo:
        return HTMLResponse("<h2>Trabajo no encontrado.</h2><a href='/'>Volver</a>", status_code=404)

    usuario_prizma = usuario_prizma.strip()
    if not usuario_prizma or not contrasena_prizma:
        return generar_html(
            resultado=trabajo.get("resultado_analisis"),
            trabajo_id=trabajo_id,
            error="Debes ingresar el usuario y la contraseña de PRIZMA.",
        )

    if trabajo.get("etapa") not in ["analizado", "error"]:
        return HTMLResponse("<h2>Este trabajo ya fue iniciado.</h2>")

    # Validar el acceso antes de encolar el curso.
    trabajo["etapa"] = "validando_login"
    trabajo["mensaje"] = "Validando usuario y contraseña en PRIZMA..."
    validacion = validar_credenciales_prizma(usuario_prizma, contrasena_prizma)

    if not validacion.get("ok"):
        trabajo["etapa"] = "analizado"
        trabajo["mensaje"] = validacion.get("mensaje", "No fue posible validar el acceso a PRIZMA.")
        return generar_html(
            resultado=trabajo.get("resultado_analisis"),
            trabajo_id=trabajo_id,
            error=trabajo["mensaje"],
        )

    _encolar_trabajo(trabajo_id, usuario_prizma, contrasena_prizma)

    html_progreso = r"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Cargue en progreso - Auto Prizma Pro</title>
        <style>
            :root {
                --fondo:#f7f8fc; --panel:#fff; --texto:#101828; --muted:#667085;
                --borde:#e5e7ef; --morado:#5548e8; --verde:#0ea968; --rojo:#ef4444;
                --sombra:0 10px 30px rgba(29,41,57,.06);
            }
            * { box-sizing:border-box; }
            body { margin:0; font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; background:var(--fondo); color:var(--texto); }
            .app { min-height:100vh; display:grid; grid-template-columns:235px 1fr; }
            .sidebar { position:sticky; top:0; height:100vh; background:#fff; border-right:1px solid var(--borde); padding:28px 20px; display:flex; flex-direction:column; }
            .marca { display:flex; align-items:center; gap:12px; margin-bottom:34px; }
            .logo { width:44px; height:44px; border-radius:13px; display:grid; place-items:center; background:linear-gradient(145deg,#6d5dfc,#4338ca); box-shadow:0 8px 20px rgba(79,70,229,.25); flex:0 0 44px; }
            .logo svg{width:29px;height:29px;overflow:visible}.logo svg path:first-child{fill:#fff}.logo svg path:last-child{fill:#c7d2fe}
            .marca strong { display:block; font-size:18px; }
            .marca span { display:block; color:var(--muted); font-size:12px; margin-top:3px; }
            .nav { display:grid; gap:8px; }
            .nav-item { padding:12px 14px; border-radius:11px; color:#475467; font-size:14px; display:flex; gap:11px; align-items:center; text-decoration:none; transition:.18s ease; }
            a.nav-item:hover{background:#f7f5ff;color:#4f46e5}.nav-item.activo{background:#f1efff;color:#4f46e5;font-weight:700}.nav-item.proximamente{opacity:.48;cursor:default}.nav-item.proximamente small{margin-left:auto;font-size:9px}
            .estado-servicio { margin-top:auto; border:1px solid var(--borde); border-radius:14px; padding:15px; background:#fff; }
            .servicio-linea { display:flex; align-items:center; gap:8px; color:#07894f; font-size:13px; font-weight:700; }
            .punto { width:8px; height:8px; border-radius:50%; background:#15b76a; }
            .servicio-mini { margin-top:13px; display:flex; justify-content:space-between; font-size:12px; color:var(--muted); }
            .chip { background:#eef4ff; color:#3538cd; padding:4px 8px; border-radius:999px; }
            .contenido { padding:28px 34px 45px; max-width:1450px; width:100%; margin:0 auto; }
            .panel { background:#fff; border:1px solid var(--borde); border-radius:16px; box-shadow:var(--sombra); }
            .cabecera-progreso { padding:25px 28px; }
            .cabecera-linea { display:flex; justify-content:space-between; align-items:center; gap:20px; }
            .cabecera-titulo { display:flex; gap:13px; align-items:center; }
            .icono-subida { width:43px; height:43px; border-radius:12px; background:#f1efff; color:#4f46e5; display:grid; place-items:center; font-size:21px; }
            .cabecera-titulo h1 { margin:0; font-size:22px; }
            .porcentaje { font-size:14px; color:var(--muted); }
            .porcentaje strong { font-size:18px; color:#4f46e5; }
            .barra { height:11px; background:#eef0f5; border-radius:999px; overflow:hidden; margin:22px 0 14px; }
            .progreso { height:100%; width:0%; border-radius:999px; background:linear-gradient(90deg,#6759f5,#4f46e5); transition:width .35s ease; }
            .estado { color:#4f46e5; font-size:14px; font-weight:600; line-height:1.45; }
            .numeros { display:grid; grid-template-columns:repeat(3,1fr); gap:18px; margin:20px 0; }
            .numero { padding:20px; display:flex; align-items:center; gap:15px; }
            .numero-icono { width:48px; height:48px; border-radius:50%; display:grid; place-items:center; font-weight:900; font-size:20px; }
            .numero.procesadas .numero-icono { background:#eef4ff; color:#4f46e5; }
            .numero.exitosas .numero-icono { background:#e9f9f1; color:#0a9b5b; }
            .numero.errores .numero-icono { background:#fff0f1; color:#e5484d; }
            .numero span { display:block; color:var(--muted); font-size:13px; }
            .numero strong { display:block; font-size:28px; margin-top:2px; }
            .zona { display:grid; grid-template-columns:minmax(0,1.55fr) minmax(280px,.75fr); gap:20px; align-items:start; }
            .lista-panel, .estado-panel { padding:23px; }
            .titulo-panel { margin:0 0 16px; font-size:17px; }
            .lista-actividades { max-height:570px; overflow:auto; border:1px solid var(--borde); border-radius:12px; }
            .actividad-progreso { display:grid; grid-template-columns:27px 1fr; gap:10px; align-items:start; padding:11px 13px; border-bottom:1px solid #eef0f4; background:#fff; }
            .actividad-progreso:last-child { border-bottom:0; }
            .actividad-icono { width:23px; height:23px; display:grid; place-items:center; font-size:15px; }
            .actividad-nombre { font-size:13px; line-height:1.4; }
            .actividad-error { color:#b42318; font-size:11px; margin-top:4px; word-break:break-word; }
            .actividad-ok { background:#fbfffd; }
            .actividad-error-fila { background:#fffafa; }
            .actividad-procesando { background:#f7f6ff; box-shadow:inset 3px 0 0 #5b48e8; }
            .estado-panel .estado-caja { background:#f8f9fc; border:1px solid var(--borde); border-radius:12px; padding:17px; margin-bottom:14px; }
            .estado-caja span { display:block; color:var(--muted); font-size:12px; }
            .estado-caja strong { display:block; margin-top:6px; font-size:14px; line-height:1.45; }
            .leyenda { display:grid; gap:12px; margin-top:22px; }
            .leyenda div { display:flex; align-items:center; gap:9px; color:#475467; font-size:13px; }
            .final { margin-top:18px; padding:18px; border-radius:12px; background:#ecfdf5; border:1px solid #a7f3d0; }
            .final h3 { margin:0 0 6px; }
            .final p { margin:0; color:#475467; font-size:13px; }
            .descarga { display:inline-flex; margin-top:14px; padding:11px 14px; border-radius:9px; background:#4f46e5; color:#fff; text-decoration:none; font-size:13px; font-weight:800; }
            .pie { margin-top:18px; color:var(--muted); font-size:12px; text-align:center; }
            @media(max-width:1050px) { .app{grid-template-columns:1fr}.sidebar{display:none}.contenido{padding:20px}.zona{grid-template-columns:1fr} }
            @media(max-width:680px) { .numeros{grid-template-columns:1fr}.contenido{padding:14px}.cabecera-linea{align-items:flex-start;flex-direction:column} }
        </style>
    </head>
    <body>
        <div class="app">
            <aside class="sidebar">
                <div class="marca">
                    <div class="logo" aria-label="Auto Prizma Pro">
                        <svg viewBox="0 0 48 48" aria-hidden="true">
                            <path d="M9 35.5 20.5 8.5c.8-1.9 3.4-1.9 4.2 0l4.1 9.6-5.4 12.7-3.1-7.4-5.2 12.1z"/>
                            <path d="M26.4 14.5 39 35.5h-8.2l-8.5-14.2z"/>
                        </svg>
                    </div>
                    <div><strong>Auto Prizma Pro</strong><span>Cargue automático</span></div>
                </div>
                <nav class="nav">
                    <a class="nav-item" href="/">⌂ <span>Inicio</span></a>
                    <a class="nav-item activo" href="/cargue-actual">⇧ <span>Cargue actual</span></a>
                    <a class="nav-item" href="/historial">◷ <span>Historial</span></a>
                    <a class="nav-item" href="/reportes">▥ <span>Reportes</span></a>
                </nav>
                <div class="estado-servicio">
                    <div class="servicio-linea"><span class="punto"></span> Servicio activo</div>
                    <div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div>
                    <div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div>
                </div>
            </aside>

            <main id="cargue-progreso" class="contenido">
                <section class="panel cabecera-progreso">
                    <div class="cabecera-linea">
                        <div class="cabecera-titulo">
                            <div class="icono-subida">⇧</div>
                            <h1>Cargue en progreso</h1>
                        </div>
                        <div class="porcentaje"><strong id="porcentaje">0%</strong> completado</div>
                    </div>
                    <div class="barra"><div id="progreso" class="progreso"></div></div>
                    <div id="estado" class="estado">Iniciando...</div>
                </section>

                <section class="numeros">
                    <div class="panel numero procesadas">
                        <div class="numero-icono">↻</div>
                        <div><span>Procesadas</span><strong id="procesadas">0</strong></div>
                    </div>
                    <div class="panel numero exitosas">
                        <div class="numero-icono">✓</div>
                        <div><span>Exitosas</span><strong id="exitosas">0</strong></div>
                    </div>
                    <div class="panel numero errores">
                        <div class="numero-icono">×</div>
                        <div><span>Errores</span><strong id="errores">0</strong></div>
                    </div>
                </section>

                <section class="zona">
                    <div class="panel lista-panel">
                        <h2 class="titulo-panel">Actividades del cargue</h2>
                        <div id="lista-actividades" class="lista-actividades">
                            <div class="actividad-progreso">Preparando lista de actividades...</div>
                        </div>
                        <div class="pie">El progreso se actualiza automáticamente.</div>
                    </div>

                    <div class="panel estado-panel">
                        <h2 class="titulo-panel">Estado del proceso</h2>
                        <div class="estado-caja">
                            <span>Actividad actual</span>
                            <strong id="actividad-actual">Preparando cargue...</strong>
                        </div>
                        <div class="estado-caja">
                            <span>Total de actividades</span>
                            <strong id="total-actividades">—</strong>
                        </div>
                        <div class="leyenda">
                            <div>⏳ Pendiente</div>
                            <div>🔄 Procesando</div>
                            <div>✅ Completada</div>
                            <div>❌ Con error</div>
                        </div>
                        <form action="/cancelar/__TRABAJO_ID__" method="post" style="margin-top:18px;">
                            <button type="submit" style="width:100%;border:1px solid #fecaca;background:#fff;color:#b42318;border-radius:9px;padding:11px 14px;cursor:pointer;font-weight:800;">Cancelar proceso</button>
                        </form>
                        <div id="final"></div>
                    </div>
                </section>
            </main>
        </div>

        <script>
            const trabajoId = "__TRABAJO_ID__";

            function escaparHtml(texto) {
                return String(texto ?? "")
                    .replaceAll("&", "&amp;")
                    .replaceAll("<", "&lt;")
                    .replaceAll(">", "&gt;")
                    .replaceAll('"', "&quot;")
                    .replaceAll("'", "&#039;");
            }

            function renderizarActividades(actividades) {
                const contenedor = document.getElementById("lista-actividades");

                if (!Array.isArray(actividades) || actividades.length === 0) {
                    contenedor.innerHTML = '<div class="actividad-progreso">Preparando lista de actividades...</div>';
                    return;
                }

                let html = "";
                let actual = "Preparando cargue...";

                for (const actividad of actividades) {
                    let icono = "⏳";
                    let clase = "";

                    if (actividad.estado === "procesando") {
                        icono = "🔄";
                        clase = "actividad-procesando";
                        actual = actividad.numero + ". " + actividad.nombre;
                    } else if (actividad.estado === "ok") {
                        icono = "✅";
                        clase = "actividad-ok";
                    } else if (actividad.estado === "error") {
                        icono = "❌";
                        clase = "actividad-error-fila";
                    }

                    html += '<div class="actividad-progreso ' + clase + '">' +
                        '<div class="actividad-icono">' + icono + '</div>' +
                        '<div><div class="actividad-nombre">' +
                        actividad.numero + '. ' + escaparHtml(actividad.nombre) + '</div>';

                    if (actividad.estado === "error" && actividad.error) {
                        html += '<div class="actividad-error">' + escaparHtml(actividad.error) + '</div>';
                    }

                    html += '</div></div>';
                }

                contenedor.innerHTML = html;
                document.getElementById("actividad-actual").innerText = actual;
            }

            async function revisarEstado() {
                try {
                    const respuesta = await fetch("/estado/" + trabajoId);
                    const datos = await respuesta.json();

                    document.getElementById("estado").innerText = datos.mensaje + (datos.posicion_cola ? ' · Posición en cola: ' + datos.posicion_cola : '');
                    document.getElementById("procesadas").innerText = datos.procesadas;
                    document.getElementById("exitosas").innerText = datos.exitosas;
                    document.getElementById("errores").innerText = datos.errores;
                    document.getElementById("total-actividades").innerText = datos.total || "—";

                    renderizarActividades(datos.detalle_actividades);

                    let porcentaje = 0;
                    if (datos.total > 0) {
                        porcentaje = (datos.procesadas / datos.total) * 100;
                    }
                    porcentaje = Math.min(100, porcentaje);

                    document.getElementById("progreso").style.width = porcentaje + "%";
                    document.getElementById("porcentaje").innerText = Math.round(porcentaje) + "%";

                    if (datos.terminado) {
                        document.getElementById("actividad-actual").innerText = "Proceso finalizado";

                        let html = '<div class="final"><h3>Proceso terminado</h3><p>' +
                            escaparHtml(datos.mensaje) + '</p>';

                        if (datos.reporte_disponible) {
                            html += '<a class="descarga" href="/reporte/' + trabajoId + '">Descargar reporte CSV</a>';
                        }

                        html += '</div>';
                        document.getElementById("final").innerHTML = html;
                        return;
                    }

                    setTimeout(revisarEstado, 1500);
                } catch (error) {
                    document.getElementById("estado").innerText = "Error consultando estado.";
                    setTimeout(revisarEstado, 3000);
                }
            }

            revisarEstado();
        </script>
    </body>
    </html>
    """

    pagina_progreso = html_progreso.replace(
        "__TRABAJO_ID__",
        trabajo_id,
    )
    trabajo["pagina_progreso"] = pagina_progreso

    return HTMLResponse(pagina_progreso)


def _pagina_sin_cargue_actual():
    return HTMLResponse(r'''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Cargue actual - Auto Prizma Pro</title>
<style>
:root{--fondo:#f7f8fc;--texto:#101828;--muted:#667085;--borde:#e5e7ef;--morado:#5548e8;--sombra:0 12px 34px rgba(29,41,57,.06)}*{box-sizing:border-box}body{margin:0;font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--fondo);color:var(--texto)}.app{min-height:100vh;display:grid;grid-template-columns:235px 1fr}.sidebar{height:100vh;background:#fff;border-right:1px solid var(--borde);padding:28px 20px;display:flex;flex-direction:column}.marca{display:flex;align-items:center;gap:12px;margin-bottom:34px}.logo{width:44px;height:44px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(145deg,#6d5dfc,#4338ca);box-shadow:0 8px 20px rgba(79,70,229,.25)}.logo svg{width:29px;height:29px}.logo svg path:first-child{fill:#fff}.logo svg path:last-child{fill:#c7d2fe}.marca strong{display:block;font-size:18px}.marca span{display:block;color:var(--muted);font-size:12px;margin-top:3px}.nav{display:grid;gap:8px}.nav-item{padding:12px 14px;border-radius:11px;color:#475467;font-size:14px;display:flex;gap:11px;align-items:center;text-decoration:none}.nav-item:hover{background:#f7f5ff;color:#4f46e5}.nav-item.activo{background:#f1efff;color:#4f46e5;font-weight:700}.estado-servicio{margin-top:auto;border:1px solid var(--borde);border-radius:14px;padding:15px}.servicio-linea{font-size:12px;font-weight:800;color:#07894f;margin-bottom:14px}.punto{width:8px;height:8px;background:#12b76a;border-radius:50%;display:inline-block;margin-right:7px}.servicio-mini{display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--muted);margin-top:10px}.chip{background:#eef2ff;color:#4f46e5;border-radius:999px;padding:4px 8px}.contenido{padding:30px 34px;display:grid;place-items:center}.vacio{width:min(680px,100%);background:#fff;border:1px solid var(--borde);border-radius:18px;box-shadow:var(--sombra);padding:54px 34px;text-align:center}.icono{width:68px;height:68px;border-radius:20px;background:#f1efff;color:#5548e8;display:grid;place-items:center;margin:0 auto 20px;font-size:28px;font-weight:800}.vacio h1{margin:0 0 10px;font-size:27px}.vacio p{margin:0 auto 24px;color:var(--muted);font-size:14px;line-height:1.6;max-width:480px}.boton{display:inline-flex;padding:12px 18px;border-radius:10px;background:linear-gradient(90deg,#5548e8,#6546e8);color:#fff;text-decoration:none;font-size:13px;font-weight:800}@media(max-width:900px){.app{grid-template-columns:1fr}.sidebar{display:none}.contenido{padding:18px}}
</style></head><body><div class="app"><aside class="sidebar"><div class="marca"><div class="logo"><svg viewBox="0 0 48 48"><path d="M9 35.5 20.5 8.5c.8-1.9 3.4-1.9 4.2 0l4.1 9.6-5.4 12.7-3.1-7.4-5.2 12.1z"/><path d="M26.4 14.5 39 35.5h-8.2l-8.5-14.2z"/></svg></div><div><strong>Auto Prizma Pro</strong><span>Automatización PRIZMA</span></div></div><nav class="nav"><a class="nav-item" href="/">⌂ <span>Inicio</span></a><a class="nav-item activo" href="/cargue-actual">⇧ <span>Cargue actual</span></a><a class="nav-item" href="/historial">◷ <span>Historial</span></a><a class="nav-item" href="/reportes">▥ <span>Reportes</span></a></nav><div class="estado-servicio"><div class="servicio-linea"><span class="punto"></span> Servicio activo</div><div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div><div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div></div></aside><main class="contenido"><section class="vacio"><div class="icono">⇧</div><h1>No hay cargues activos</h1><p>Cuando alguien inicie un proceso desde Inicio, aparecerá aquí para que el equipo pueda consultar su progreso.</p><a class="boton" href="/">Iniciar un nuevo cargue</a></section></main></div></body></html>''')


def _formatear_inicio_cargue(fecha_iso):
    if not fecha_iso:
        return "Iniciando..."
    try:
        fecha = datetime.fromisoformat(fecha_iso)
        if fecha.tzinfo is None:
            fecha = fecha.replace(tzinfo=ZONA_HORARIA_COLOMBIA)
        else:
            fecha = fecha.astimezone(ZONA_HORARIA_COLOMBIA)
        return fecha.strftime("%d/%m/%Y · %I:%M %p").replace("AM", "a. m.").replace("PM", "p. m.")
    except Exception:
        return "En progreso"


def _pagina_cargues_activos(trabajos):
    tarjetas = []
    for trabajo in trabajos:
        cursos = trabajo.get("cursos") or []
        if cursos:
            curso = cursos[0].get("curso") or "Curso sin nombre"
            programa = cursos[0].get("programa") or "Programa sin nombre"
            if len(cursos) > 1:
                curso = f"{curso} y {len(cursos) - 1} curso(s) más"
        else:
            curso, programa = "Curso en proceso", "Programa sin nombre"
        total = int(trabajo.get("total") or 0)
        procesadas = int(trabajo.get("procesadas") or 0)
        exitosas = int(trabajo.get("exitosas") or 0)
        errores = int(trabajo.get("errores") or 0)
        porcentaje = max(0, min(100, round((procesadas / total) * 100) if total > 0 else 0))
        trabajo_id = html.escape(str(trabajo.get("id") or ""))
        en_cola = trabajo.get("etapa") == "en_cola"
        posicion = _posicion_en_cola(str(trabajo.get("id") or "")) if en_cola else None
        estado_texto = f"En cola · #{posicion}" if posicion else "En progreso"
        inicio_label = "En cola desde" if en_cola else "Iniciado"
        fecha_mostrar = trabajo.get("encolado_en") if en_cola else trabajo.get("iniciado_en")
        tarjetas.append(f'''<article class="carga-card"><div class="curso-bloque"><div class="curso-icono">▣</div><div class="curso-texto"><h2>{html.escape(curso)}</h2><p>Programa: {html.escape(programa)}</p><span class="estado-chip"><span></span> {html.escape(estado_texto)}</span></div></div><div class="inicio-bloque"><small>◷ {inicio_label}</small><strong>{html.escape(_formatear_inicio_cargue(fecha_mostrar))}</strong></div><div class="avance-bloque"><div class="avance-cab"><span>Progreso general</span><strong>{porcentaje}%</strong></div><div class="barra"><div class="barra-interna" style="width:{porcentaje}%"></div></div><div class="metricas"><div><i class="verde"></i><small>Procesadas</small><strong>{procesadas}</strong></div><div><i class="verde"></i><small>Exitosas</small><strong>{exitosas}</strong></div><div><i class="rojo"></i><small>Errores</small><strong>{errores}</strong></div></div></div><div class="accion-bloque"><a href="/proceso/{trabajo_id}">Entrar al proceso <b>›</b></a><form action="/cancelar/{trabajo_id}" method="post" style="margin-top:8px"><button type="submit" style="border:1px solid #fecaca;background:#fff;color:#b42318;border-radius:8px;padding:8px 10px;cursor:pointer;font-weight:700">Cancelar</button></form></div></article>''')
    cuerpo = "".join(tarjetas)
    cantidad = len(trabajos)
    plural = "s" if cantidad != 1 else ""
    return HTMLResponse(f'''<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Cargue actual - Auto Prizma Pro</title><style>
:root{{--fondo:#f7f8fc;--texto:#101828;--muted:#667085;--borde:#e5e7ef;--morado:#5548e8;--sombra:0 10px 30px rgba(29,41,57,.05)}}*{{box-sizing:border-box}}body{{margin:0;font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--fondo);color:var(--texto)}}.app{{min-height:100vh;display:grid;grid-template-columns:235px 1fr}}.sidebar{{position:sticky;top:0;height:100vh;background:#fff;border-right:1px solid var(--borde);padding:28px 20px;display:flex;flex-direction:column}}.marca{{display:flex;align-items:center;gap:12px;margin-bottom:34px}}.logo{{width:44px;height:44px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(145deg,#6d5dfc,#4338ca);box-shadow:0 8px 20px rgba(79,70,229,.25)}}.logo svg{{width:29px;height:29px}}.logo svg path:first-child{{fill:#fff}}.logo svg path:last-child{{fill:#c7d2fe}}.marca strong{{display:block;font-size:18px}}.marca span{{display:block;color:var(--muted);font-size:12px;margin-top:3px}}.nav{{display:grid;gap:8px}}.nav-item{{padding:12px 14px;border-radius:11px;color:#475467;font-size:14px;display:flex;gap:11px;align-items:center;text-decoration:none}}.nav-item:hover{{background:#f7f5ff;color:#4f46e5}}.nav-item.activo{{background:#f1efff;color:#4f46e5;font-weight:700}}.estado-servicio{{margin-top:auto;border:1px solid var(--borde);border-radius:14px;padding:15px}}.servicio-linea{{font-size:12px;font-weight:800;color:#07894f;margin-bottom:14px}}.punto{{width:8px;height:8px;background:#12b76a;border-radius:50%;display:inline-block;margin-right:7px}}.servicio-mini{{display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--muted);margin-top:10px}}.chip{{background:#eef2ff;color:#4f46e5;border-radius:999px;padding:4px 8px}}.contenido{{padding:34px 36px 44px;min-width:0}}.cabecera-top{{display:flex;justify-content:space-between;gap:20px;align-items:flex-start;margin-bottom:24px}}.cabecera h1{{font-size:29px;margin:0 0 7px}}.cabecera p{{margin:0;color:var(--muted);font-size:14px;line-height:1.5}}.contador{{background:#f1efff;color:#5548e8;border-radius:999px;padding:8px 12px;font-size:12px;font-weight:800;white-space:nowrap}}.lista{{display:grid;gap:16px}}.carga-card{{background:#fff;border:1px solid var(--borde);border-radius:17px;box-shadow:var(--sombra);padding:22px;display:grid;grid-template-columns:minmax(245px,1.15fr) 170px minmax(320px,1.25fr) 185px;gap:22px;align-items:center}}.curso-bloque{{display:flex;gap:14px;align-items:center;min-width:0}}.curso-icono{{width:58px;height:58px;border-radius:15px;background:#f1efff;color:#5548e8;display:grid;place-items:center;font-size:24px;flex:0 0 58px}}.curso-texto{{min-width:0}}.curso-texto h2{{font-size:16px;margin:0 0 5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.curso-texto p{{font-size:12px;color:var(--muted);margin:0 0 10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.estado-chip{{display:inline-flex;align-items:center;gap:6px;background:#eef6ff;color:#1473e6;border-radius:999px;padding:5px 9px;font-size:11px;font-weight:800}}.estado-chip span{{width:6px;height:6px;background:#2e90fa;border-radius:50%}}.inicio-bloque small{{display:block;color:var(--muted);font-size:11px;margin-bottom:7px}}.inicio-bloque strong{{font-size:12px;color:#475467}}.avance-cab{{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;color:var(--muted);font-size:11px}}.avance-cab strong{{font-size:21px;color:#5548e8}}.barra{{height:10px;border-radius:999px;background:#eeecff;overflow:hidden}}.barra-interna{{height:100%;background:linear-gradient(90deg,#6759f5,#4f46e5);border-radius:999px}}.metricas{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:12px}}.metricas div{{display:grid;grid-template-columns:8px 1fr;column-gap:5px;align-items:center}}.metricas i{{width:7px;height:7px;border-radius:50%}}.metricas i.verde{{background:#12b76a}}.metricas i.rojo{{background:#ef4444}}.metricas small{{font-size:10px;color:var(--muted)}}.metricas strong{{grid-column:2;font-size:14px;margin-top:2px}}.accion-bloque{{border-left:1px solid var(--borde);padding-left:20px;text-align:center}}.accion-bloque a{{display:flex;justify-content:center;align-items:center;gap:8px;background:linear-gradient(90deg,#5548e8,#6546e8);color:#fff;text-decoration:none;font-size:12px;font-weight:800;border-radius:10px;padding:12px 13px}}.accion-bloque b{{font-size:18px}}.accion-bloque small{{display:block;color:var(--muted);font-size:10px;margin-top:8px}}.nota{{margin-top:18px;border:1px solid #ddd8ff;background:#faf9ff;border-radius:14px;padding:14px 17px;color:#475467;font-size:12px}}@media(max-width:1200px){{.carga-card{{grid-template-columns:1fr 1fr}}.accion-bloque{{border-left:0;padding-left:0}}}}@media(max-width:900px){{.app{{grid-template-columns:1fr}}.sidebar{{display:none}}.contenido{{padding:20px}}.carga-card{{grid-template-columns:1fr}}}}
</style></head><body><div class="app"><aside class="sidebar"><div class="marca"><div class="logo"><svg viewBox="0 0 48 48"><path d="M9 35.5 20.5 8.5c.8-1.9 3.4-1.9 4.2 0l4.1 9.6-5.4 12.7-3.1-7.4-5.2 12.1z"/><path d="M26.4 14.5 39 35.5h-8.2l-8.5-14.2z"/></svg></div><div><strong>Auto Prizma Pro</strong><span>Automatización PRIZMA</span></div></div><nav class="nav"><a class="nav-item" href="/">⌂ <span>Inicio</span></a><a class="nav-item activo" href="/cargue-actual">⇧ <span>Cargue actual</span></a><a class="nav-item" href="/historial">◷ <span>Historial</span></a><a class="nav-item" href="/reportes">▥ <span>Reportes</span></a></nav><div class="estado-servicio"><div class="servicio-linea"><span class="punto"></span> Servicio activo</div><div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div><div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div></div></aside><main class="contenido"><section class="cabecera"><div class="cabecera-top"><div><h1>Cargue actual</h1><p>Aquí puedes ver todos los cursos que están siendo procesados actualmente.<br>Entra a un proceso para revisar el detalle de sus actividades.</p></div><div class="contador">{cantidad} proceso{plural} activo{plural}</div></div></section><section class="lista">{cuerpo}</section><div class="nota"><strong>Vista global del equipo.</strong> Cada tarjeta representa un cargue activo. Para ver la lista detallada de actividades, entra al proceso correspondiente.</div></main></div><script>setTimeout(function(){{window.location.reload();}},3000);</script></body></html>''')


@app.get("/cargue-actual", response_class=HTMLResponse)
def cargue_actual():
    activos = [
        trabajo for trabajo in TRABAJOS.values()
        if not trabajo.get("terminado")
        and trabajo.get("etapa") in {"en_cola", "iniciando", "login", "preparando", "procesando", "validando_login"}
    ]
    activos.sort(key=lambda item: item.get("encolado_en") or item.get("iniciado_en") or item.get("creado_en", ""))
    if not activos:
        return _pagina_sin_cargue_actual()
    return _pagina_cargues_activos(activos)


@app.post("/cancelar/{trabajo_id}", response_class=HTMLResponse)
def cancelar_trabajo(trabajo_id: str):
    with COLA_CONDICION:
        trabajo = TRABAJOS.get(trabajo_id)
        if not trabajo:
            return HTMLResponse("<h2>Proceso no encontrado.</h2>", status_code=404)

        if trabajo.get("terminado"):
            return HTMLResponse(
                "<script>window.location='/cargue-actual';</script>"
            )

        trabajo["cancelar_solicitado"] = True

        if trabajo.get("etapa") == "en_cola":
            try:
                COLA_CARGUES.remove(trabajo_id)
            except ValueError:
                pass
            trabajo.pop("usuario_prizma_temporal", None)
            trabajo.pop("contrasena_prizma_temporal", None)
            trabajo["etapa"] = "cancelado"
            trabajo["mensaje"] = "Proceso cancelado antes de iniciar."
            trabajo["terminado"] = True
            trabajo["finalizado_en"] = _ahora_colombia_iso()
        else:
            trabajo["mensaje"] = (
                "Cancelación solicitada. El proceso se detendrá de forma segura "
                "al terminar la actividad actual."
            )

        COLA_CONDICION.notify_all()

    return HTMLResponse("<script>window.location='/cargue-actual';</script>")


@app.get("/proceso/{trabajo_id}", response_class=HTMLResponse)
def ver_proceso(trabajo_id: str):
    trabajo = TRABAJOS.get(trabajo_id)
    if not trabajo:
        return HTMLResponse("<h2>Proceso no encontrado.</h2><p><a href='/cargue-actual'>Volver a Cargue actual</a></p>", status_code=404)
    pagina = trabajo.get("pagina_progreso")
    if pagina:
        return HTMLResponse(pagina)
    return HTMLResponse("<h2>El proceso todavía se está preparando.</h2><p><a href='/cargue-actual'>Volver a Cargue actual</a></p>", status_code=202)


# ============================================================
# ESTADO
# ============================================================

@app.get(
    "/estado/{trabajo_id}",
)
def estado_trabajo(
    trabajo_id: str,
):

    trabajo = TRABAJOS.get(
        trabajo_id
    )

    if not trabajo:

        return JSONResponse(
            {
                "error":
                    "Trabajo no encontrado"
            },
            status_code=404,
        )

    return {
        "etapa":
            trabajo["etapa"],

        "mensaje":
            trabajo["mensaje"],

        "total":
            trabajo["total"],

        "procesadas":
            trabajo["procesadas"],

        "exitosas":
            trabajo["exitosas"],

        "errores":
            trabajo["errores"],

        "terminado":
            trabajo["terminado"],

        "detalle_actividades":
            trabajo.get(
                "detalle_actividades",
                [],
            ),

        "reporte_disponible":
            os.path.isfile(
                trabajo["ruta_reporte"]
            ),
        "posicion_cola": _posicion_en_cola(trabajo_id),
        "cancelar_solicitado": bool(trabajo.get("cancelar_solicitado")),
    }


# ============================================================
# HISTORIAL Y REPORTES
# ============================================================

def _formatear_fecha_hora(fecha_iso):
    try:
        fecha = datetime.fromisoformat(fecha_iso)
        if fecha.tzinfo is None:
            fecha = fecha.replace(tzinfo=ZONA_HORARIA_COLOMBIA)
        else:
            fecha = fecha.astimezone(ZONA_HORARIA_COLOMBIA)
    except Exception:
        return "—", "—"

    fecha_texto = fecha.strftime("%d/%m/%Y")
    hora = fecha.strftime("%I:%M").lstrip("0") or "0:00"
    sufijo = "a. m." if fecha.hour < 12 else "p. m."
    return fecha_texto, f"{hora} {sufijo}"


def _pagina_registros(tipo="historial"):
    registros = list(reversed(_cargar_historial()))
    es_historial = tipo == "historial"
    activo_historial = "activo" if es_historial else ""
    activo_reportes = "" if es_historial else "activo"

    filas = []
    programas = set()

    if es_historial:
        for registro in registros:
            fecha, hora = _formatear_fecha_hora(registro.get("fecha_iso", ""))
            cursos = registro.get("cursos") or [{"curso": "Curso sin nombre", "programa": "Programa sin nombre"}]

            for item in cursos:
                curso = str(item.get("curso") or "Curso sin nombre")
                programa = str(item.get("programa") or "Programa sin nombre")
                programas.add(programa)
                busqueda = html.escape(f"{curso} {programa}".lower(), quote=True)
                programa_attr = html.escape(programa, quote=True)
                registro_id = html.escape(str(registro.get("id") or ""), quote=True)

                filas.append(
                    f'<tr class="fila-registro" data-busqueda="{busqueda}" data-programa="{programa_attr}">'
                    f'<td class="fecha-celda"><span class="icono-fecha">▣</span><div><strong>{html.escape(fecha)}</strong><small>{html.escape(hora)}</small></div></td>'
                    f'<td>{html.escape(curso)}</td>'
                    f'<td>{html.escape(programa)}</td>'
                    f'<td class="accion-celda"><a class="boton-reporte" href="/reporte-guardado/{registro_id}">◉ <span>Ver reporte</span></a></td>'
                    '</tr>'
                )
    else:
        for registro in registros:
            fecha, hora = _formatear_fecha_hora(registro.get("fecha_iso", ""))
            cursos = registro.get("cursos") or [{"curso": "Curso sin nombre", "programa": "Programa sin nombre"}]
            principal = cursos[0]
            curso = str(principal.get("curso") or "Curso sin nombre")
            programa = str(principal.get("programa") or "Programa sin nombre")
            archivo = str(registro.get("archivo_reporte") or "reporte.csv")
            programas.add(programa)
            busqueda = html.escape(f"{archivo} {curso} {programa}".lower(), quote=True)
            programa_attr = html.escape(programa, quote=True)
            registro_id = html.escape(str(registro.get("id") or ""), quote=True)

            filas.append(
                f'<tr class="fila-registro" data-busqueda="{busqueda}" data-programa="{programa_attr}">'
                f'<td class="fecha-celda"><span class="icono-fecha">▣</span><div><strong>{html.escape(fecha)}</strong><small>{html.escape(hora)}</small></div></td>'
                f'<td class="archivo-reporte" title="{html.escape(archivo, quote=True)}">{html.escape(archivo)}</td>'
                f'<td>{html.escape(curso)}</td>'
                f'<td>{html.escape(programa)}</td>'
                f'<td class="accion-celda"><a class="boton-reporte" href="/reporte-guardado/{registro_id}">⇩ <span>Descargar</span></a></td>'
                '</tr>'
            )

    opciones_programa = ''.join(
        f'<option value="{html.escape(p, quote=True)}">{html.escape(p)}</option>'
        for p in sorted(programas, key=str.casefold)
    )

    if es_historial:
        titulo = "Historial de cargues"
        subtitulo = "Consulta los cursos procesados anteriormente y accede a su reporte correspondiente."
        contador = sum(len(r.get("cursos") or [1]) for r in registros)
        contador_texto = f"{contador} registros"
        placeholder = "Buscar curso o programa..."
        tabla_cabecera = "<th>Fecha y hora</th><th>Curso</th><th>Programa</th><th>Reporte</th>"
        min_table = "900px"
        vacio = "Todavía no hay cargues registrados."
    else:
        titulo = "Reportes generados"
        subtitulo = "Todos los reportes se guardan automáticamente con el nombre del curso y el programa correspondiente."
        contador = len(registros)
        contador_texto = f"{contador} reportes"
        placeholder = "Buscar reporte, curso o programa..."
        tabla_cabecera = "<th>Fecha</th><th>Nombre del reporte</th><th>Curso</th><th>Programa</th><th>Acción</th>"
        min_table = "1120px"
        vacio = "Todavía no hay reportes guardados."

    filas_html = ''.join(filas)
    if not filas_html:
        filas_html = f'<tr><td colspan="5" class="vacio">{html.escape(vacio)}</td></tr>'

    pagina = '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>__TITULO__ - Auto Prizma Pro</title>
    <style>
        :root{--fondo:#f7f8fc;--panel:#fff;--texto:#101828;--muted:#667085;--borde:#e5e7ef;--morado:#5548e8;--sombra:0 12px 34px rgba(29,41,57,.06)}
        *{box-sizing:border-box} body{margin:0;font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--fondo);color:var(--texto)}
        .app{min-height:100vh;display:grid;grid-template-columns:235px 1fr}.sidebar{position:sticky;top:0;height:100vh;background:#fff;border-right:1px solid var(--borde);padding:28px 20px;display:flex;flex-direction:column}
        .marca{display:flex;align-items:center;gap:12px;margin-bottom:34px}.logo{width:44px;height:44px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(145deg,#6d5dfc,#4338ca);box-shadow:0 8px 20px rgba(79,70,229,.25)}
        .logo svg{width:29px;height:29px;overflow:visible}.logo svg path:first-child{fill:#fff}.logo svg path:last-child{fill:#c7d2fe}.marca strong{display:block;font-size:18px}.marca span{display:block;color:var(--muted);font-size:12px;margin-top:3px}
        .nav{display:grid;gap:8px}.nav-item{padding:12px 14px;border-radius:11px;color:#475467;font-size:14px;display:flex;gap:11px;align-items:center;text-decoration:none;transition:.18s ease}.nav-item:hover{background:#f7f5ff;color:#4f46e5}.nav-item.activo{background:#f1efff;color:#4f46e5;font-weight:700}
        .estado-servicio{margin-top:auto;border:1px solid var(--borde);border-radius:14px;padding:15px}.servicio-linea{font-size:12px;font-weight:800;color:#07894f;margin-bottom:14px}.punto{width:8px;height:8px;background:#12b76a;border-radius:50%;display:inline-block;margin-right:7px}.servicio-mini{display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--muted);margin-top:10px}.chip{background:#eef2ff;color:#4f46e5;border-radius:999px;padding:4px 8px}
        .contenido{padding:30px 34px}.panel{background:#fff;border:1px solid var(--borde);border-radius:18px;box-shadow:var(--sombra);padding:28px;max-width:1320px;margin:0 auto}.cabecera{display:flex;align-items:flex-start;justify-content:space-between;gap:20px}h1{font-size:28px;margin:0 0 7px}.subtitulo{margin:0;color:var(--muted);font-size:14px}.contador{background:#f1efff;color:#4f46e5;border-radius:10px;padding:9px 12px;font-size:12px;font-weight:800;white-space:nowrap}
        .herramientas{display:grid;grid-template-columns:minmax(260px,1fr) 290px;gap:18px;margin:28px 0 22px}.buscador,.filtro{height:52px;border:1px solid #d8dce6;border-radius:11px;background:#fff;display:flex;align-items:center;gap:10px;padding:0 15px;color:#667085}.buscador input,.filtro select{width:100%;border:0;outline:0;background:transparent;font-size:14px;color:#475467}.filtro select{cursor:pointer}
        .tabla-wrap{border:1px solid var(--borde);border-radius:14px;overflow:auto}table{width:100%;border-collapse:collapse;font-size:13px;min-width:__MIN_TABLE__}thead{background:#faf9ff}th{text-align:left;padding:15px 16px;color:#4338ca;font-size:12px;border-bottom:1px solid var(--borde)}td{padding:15px 16px;border-bottom:1px solid #eef0f4;vertical-align:middle;color:#344054}tbody tr:last-child td{border-bottom:0}.fecha-celda{display:flex;align-items:center;gap:11px;white-space:nowrap}.fecha-celda strong{display:block;font-weight:500}.fecha-celda small{display:block;color:#98a2b3;margin-top:4px}.icono-fecha{width:34px;height:34px;border-radius:9px;background:#f4f1ff;color:#6d5dfc;display:grid;place-items:center}.accion-celda{white-space:nowrap}.boton-reporte{display:inline-flex;align-items:center;gap:7px;border:1px solid #d7d0ff;color:#5b4ce8;text-decoration:none;border-radius:9px;padding:9px 12px;font-weight:700;font-size:12px;background:#fff}.boton-reporte:hover{background:#f7f5ff}.archivo-reporte{max-width:390px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vacio{text-align:center;color:#98a2b3;padding:50px 20px!important}.oculta{display:none}
        @media(max-width:1050px){.app{grid-template-columns:1fr}.sidebar{display:none}.contenido{padding:20px}}@media(max-width:700px){.contenido{padding:12px}.panel{padding:18px}.cabecera{flex-direction:column}.herramientas{grid-template-columns:1fr}}
    </style>
</head>
<body>
<div class="app">
    <aside class="sidebar">
        <div class="marca"><div class="logo"><svg viewBox="0 0 48 48" aria-hidden="true"><path d="M9 35.5 20.5 8.5c.8-1.9 3.4-1.9 4.2 0l4.1 9.6-5.4 12.7-3.1-7.4-5.2 12.1z"/><path d="M26.4 14.5 39 35.5h-8.2l-8.5-14.2z"/></svg></div><div><strong>Auto Prizma Pro</strong><span>Automatización PRIZMA</span></div></div>
        <nav class="nav">
            <a class="nav-item" href="/">⌂ <span>Inicio</span></a>
            <a class="nav-item" href="/cargue-actual">⇧ <span>Cargue actual</span></a>
            <a class="nav-item __ACTIVO_HISTORIAL__" href="/historial">◷ <span>Historial</span></a>
            <a class="nav-item __ACTIVO_REPORTES__" href="/reportes">▥ <span>Reportes</span></a>
        </nav>
        <div class="estado-servicio"><div class="servicio-linea"><span class="punto"></span> Servicio activo</div><div class="servicio-mini"><span>Navegador</span><span class="chip">Chromium</span></div><div class="servicio-mini"><span>Conexión</span><span class="chip">Estable</span></div></div>
    </aside>
    <main class="contenido"><section class="panel"><div class="cabecera"><div><h1>__TITULO__</h1><p class="subtitulo">__SUBTITULO__</p></div><div class="contador">▣ &nbsp;__CONTADOR__</div></div><div class="herramientas"><label class="buscador">⌕ <input id="buscar" type="search" placeholder="__PLACEHOLDER__"></label><label class="filtro">▽ <select id="programa"><option value="">Todos los programas</option>__OPCIONES__</select></label></div><div class="tabla-wrap"><table><thead><tr>__CABECERA__</tr></thead><tbody>__FILAS__</tbody></table></div></section></main>
</div>
<script>
const buscar=document.getElementById('buscar');const programa=document.getElementById('programa');function filtrar(){const texto=(buscar.value||'').trim().toLowerCase();const p=programa.value||'';document.querySelectorAll('.fila-registro').forEach(f=>{const okTexto=!texto||(f.dataset.busqueda||'').includes(texto);const okPrograma=!p||f.dataset.programa===p;f.classList.toggle('oculta',!(okTexto&&okPrograma));});}buscar.addEventListener('input',filtrar);programa.addEventListener('change',filtrar);
</script>
</body></html>'''

    pagina = (pagina
        .replace("__TITULO__", html.escape(titulo))
        .replace("__SUBTITULO__", html.escape(subtitulo))
        .replace("__CONTADOR__", html.escape(contador_texto))
        .replace("__PLACEHOLDER__", html.escape(placeholder, quote=True))
        .replace("__OPCIONES__", opciones_programa)
        .replace("__CABECERA__", tabla_cabecera)
        .replace("__FILAS__", filas_html)
        .replace("__MIN_TABLE__", min_table)
        .replace("__ACTIVO_HISTORIAL__", activo_historial)
        .replace("__ACTIVO_REPORTES__", activo_reportes)
    )
    return HTMLResponse(pagina)


@app.get("/historial", response_class=HTMLResponse)
def historial_cargues():
    return _pagina_registros("historial")


@app.get("/reportes", response_class=HTMLResponse)
def reportes_generados():
    return _pagina_registros("reportes")


@app.get("/reporte-guardado/{registro_id}")
def descargar_reporte_guardado(registro_id: str):
    registro = next((r for r in _cargar_historial() if r.get("id") == registro_id), None)

    if not registro:
        return JSONResponse({"error": "Reporte no encontrado"}, status_code=404)

    nombre = os.path.basename(str(registro.get("archivo_reporte") or ""))
    ruta = os.path.join(RESULTADOS_DIR, nombre)

    if not nombre or not os.path.isfile(ruta):
        return JSONResponse({"error": "El archivo del reporte ya no está disponible"}, status_code=404)

    return FileResponse(ruta, media_type="text/csv", filename=nombre)


# ============================================================
# DESCARGAR REPORTE
# ============================================================

@app.get(
    "/reporte/{trabajo_id}",
)
def descargar_reporte(
    trabajo_id: str,
):

    trabajo = TRABAJOS.get(
        trabajo_id
    )

    if not trabajo:

        return JSONResponse(
            {
                "error":
                    "Trabajo no encontrado"
            },
            status_code=404,
        )

    ruta = trabajo[
        "ruta_reporte"
    ]

    if not os.path.isfile(
        ruta
    ):

        return JSONResponse(
            {
                "error":
                    "Reporte todavía no disponible"
            },
            status_code=404,
        )

    return FileResponse(
        ruta,
        media_type="text/csv",
        filename=os.path.basename(ruta),
    )