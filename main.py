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

from openpyxl import load_workbook

from motor_prizma import (
    ejecutar_cargue,
    determinar_tipo_archivo,
    normalizar_categoria,
    normalizar_texto,
)

import zipfile
import os
import uuid


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

RESULTADOS_DIR = os.path.join(
    BASE_DIR,
    "resultados",
)


for carpeta in [
    UPLOADS_DIR,
    TEMP_DIR,
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


# ============================================================
# ANALIZAR EXCEL
# ============================================================

def analizar_excel(
    contenido_excel,
    procesar_ovi=True,
    procesar_ova=True,
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

    bloque = ""

    # ========================================================
    # ERROR
    # ========================================================

    if error:

        bloque = f"""
        <div class="error">

            <strong>
                ERROR
            </strong>

            <p>
                {error}
            </p>

        </div>
        """

    # ========================================================
    # RESULTADO
    # ========================================================

    elif resultado:

        hojas_html = ""

        total_actividades = 0

        for hoja in resultado[
            "hojas"
        ]:

            filas = ""

            total_actividades += len(
                hoja["actividades"]
            )

            for actividad in hoja[
                "actividades"
            ]:

                filas += f"""
                <tr>

                    <td>
                        {actividad["fila"]}
                    </td>

                    <td>
                        {actividad["semana"]}
                    </td>

                    <td>
                        {actividad["unidad"]}
                    </td>

                    <td>
                        {actividad["nombre"]}
                    </td>

                    <td>
                        {actividad["categoria"]}
                    </td>

                    <td>
                        {actividad["tipo"]}
                    </td>

                </tr>
                """

            hojas_html += f"""
            <div class="panel">

                <h2>
                    {hoja["curso"]}
                </h2>

                <div class="grid">

                    <div class="dato">

                        <span>
                            Programa
                        </span>

                        <strong>
                            {hoja["programa"]}
                        </strong>

                    </div>

                    <div class="dato">

                        <span>
                            OVI
                        </span>

                        <strong>
                            {hoja["ovi"]}
                        </strong>

                    </div>

                    <div class="dato">

                        <span>
                            OVA
                        </span>

                        <strong>
                            {hoja["ova"]}
                        </strong>

                    </div>

                    <div class="dato">

                        <span>
                            H5P
                        </span>

                        <strong>
                            {hoja["h5p"]}
                        </strong>

                    </div>

                    <div class="dato">

                        <span>
                            PDF
                        </span>

                        <strong>
                            {hoja["pdf"]}
                        </strong>

                    </div>

                </div>

                <h3>
                    Actividades a procesar
                </h3>

                <div class="tabla">

                    <table>

                        <thead>

                            <tr>

                                <th>
                                    Fila
                                </th>

                                <th>
                                    Semana
                                </th>

                                <th>
                                    Unidad
                                </th>

                                <th>
                                    Actividad
                                </th>

                                <th>
                                    Categoría
                                </th>

                                <th>
                                    Tipo
                                </th>

                            </tr>

                        </thead>

                        <tbody>
                            {filas}
                        </tbody>

                    </table>

                </div>

            </div>
            """

        zip_info = resultado[
            "zip"
        ]

        bloque = f"""

        <div class="panel correcto">

            <h2>
                ✅ Análisis completado
            </h2>

            <div class="grid">

                <div class="dato">

                    <span>
                        Actividades a procesar
                    </span>

                    <strong>
                        {total_actividades}
                    </strong>

                </div>

                <div class="dato">

                    <span>
                        H5P en ZIP
                    </span>

                    <strong>
                        {zip_info["h5p"]}
                    </strong>

                </div>

                <div class="dato">

                    <span>
                        PDF en ZIP
                    </span>

                    <strong>
                        {zip_info["pdf"]}
                    </strong>

                </div>

                <div class="dato">

                    <span>
                        Total recursos
                    </span>

                    <strong>
                        {zip_info["total"]}
                    </strong>

                </div>

            </div>

        </div>

        {hojas_html}

        <div class="panel iniciar">

            <h2>
                Iniciar cargue
            </h2>

            <p>
                Se procesarán las
                <strong>
                    {total_actividades}
                    actividades compatibles
                </strong>
                mostradas arriba.
            </p>

            <div class="reglas">

                <div>
                    ✅ OVI y OVA
                </div>

                <div>
                    ✅ H5P y PDF se cargan en Recurso
                </div>

                <div>
                    🚫 Material descargable no se toca
                </div>

                <div>
                    🚫 Retos Evaluativos excluidos
                </div>

                <div>
                    🚫 Video Intro/Cierre excluidos
                </div>

                <div>
                    ✅ Error individual no detiene el curso
                </div>

            </div>

            <form
                action="/iniciar/{trabajo_id}"
                method="post"
            >

                <button
                    class="verde"
                    type="submit"
                >

                    INICIAR CARGUE COMPLETO EN PRIZMA

                </button>

            </form>

        </div>

        """

    return f"""
    <!DOCTYPE html>

    <html lang="es">

    <head>

        <meta charset="UTF-8">

        <meta
            name="viewport"
            content="width=device-width,
            initial-scale=1.0"
        >

        <title>
            Auto Prizma Pro
        </title>

        <style>

            * {{
                box-sizing: border-box;
            }}

            body {{
                margin: 0;
                background: #f4f6fb;
                color: #1f2937;
                font-family: Arial, sans-serif;
            }}

            .contenedor {{
                width: 1100px;
                max-width: calc(100% - 40px);
                margin: 45px auto;
            }}

            h1 {{
                margin-bottom: 6px;
            }}

            .subtitulo {{
                color: #6b7280;
                margin-bottom: 30px;
            }}

            .panel {{
                background: white;
                padding: 28px;
                margin-top: 25px;
                border-radius: 16px;
                box-shadow:
                    0 8px 30px
                    rgba(0,0,0,.06);
            }}

            .principal {{
                margin-top: 0;
            }}

            .campo {{
                margin-bottom: 22px;
            }}

            label {{
                display: block;
                font-weight: bold;
                margin-bottom: 8px;
            }}

            input[type=file] {{
                width: 100%;
                padding: 13px;
                border:
                    1px solid #d1d5db;
                border-radius: 10px;
            }}

            .checks {{
                display: flex;
                gap: 24px;
                margin-bottom: 22px;
            }}

            .checks label {{
                display: inline;
                font-weight: normal;
            }}

            .aviso {{
                padding: 14px;
                border-radius: 10px;
                background: #fff7ed;
                margin-bottom: 22px;
            }}

            button {{
                width: 100%;
                padding: 15px;
                border: 0;
                border-radius: 10px;
                background: #111827;
                color: white;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
            }}

            button:hover {{
                opacity: .92;
            }}

            .verde {{
                background: #16a34a;
            }}

            .correcto {{
                background: #ecfdf5;
            }}

            .iniciar {{
                border: 2px solid #16a34a;
            }}

            .error {{
                background: #fef2f2;
                padding: 20px;
                margin-top: 25px;
                border-radius: 12px;
            }}

            .grid {{
                display: grid;
                grid-template-columns:
                    repeat(
                        auto-fit,
                        minmax(160px, 1fr)
                    );
                gap: 12px;
                margin: 18px 0;
            }}

            .dato {{
                padding: 14px;
                border:
                    1px solid #e5e7eb;
                border-radius: 10px;
                background: #f8fafc;
            }}

            .dato span {{
                display: block;
                color: #6b7280;
                font-size: 13px;
                margin-bottom: 5px;
            }}

            .tabla {{
                overflow-x: auto;
            }}

            table {{
                width: 100%;
                border-collapse: collapse;
            }}

            th {{
                text-align: left;
                background: #f3f4f6;
                padding: 11px;
            }}

            td {{
                padding: 11px;
                border-bottom:
                    1px solid #e5e7eb;
            }}

            .reglas {{
                background: #f8fafc;
                border-radius: 10px;
                padding: 15px;
                margin: 20px 0;
                line-height: 1.9;
            }}

        </style>

    </head>

    <body>

        <div class="contenedor">

            <h1>
                Auto Prizma Pro
            </h1>

            <div class="subtitulo">
                Automatización de cargue
                de actividades PRIZMA
            </div>

            <div class="panel principal">

                <form
                    action="/analizar"
                    method="post"
                    enctype="multipart/form-data"
                >

                    <div class="campo">

                        <label>
                            Archivo Excel
                        </label>

                        <input
                            type="file"
                            name="excel"
                            accept=".xlsx"
                            required
                        >

                    </div>

                    <div class="campo">

                        <label>
                            ZIP de recursos
                        </label>

                        <input
                            type="file"
                            name="recursos"
                            accept=".zip"
                            required
                        >

                    </div>

                    <div class="checks">

                        <div>

                            <input
                                id="ovi"
                                type="checkbox"
                                name="ovi"
                                value="1"
                                checked
                            >

                            <label for="ovi">
                                OVI
                            </label>

                        </div>

                        <div>

                            <input
                                id="ova"
                                type="checkbox"
                                name="ova"
                                value="1"
                                checked
                            >

                            <label for="ova">
                                OVA
                            </label>

                        </div>

                    </div>

                    <div class="aviso">

                        Retos Evaluativos,
                        Video Intro y Video Cierre
                        están excluidos.

                    </div>

                    <button type="submit">

                        Analizar archivos

                    </button>

                </form>

            </div>

            {bloque}

        </div>

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
):

    try:

        if not excel.filename.lower().endswith(
            ".xlsx"
        ):

            return generar_html(
                error=(
                    "El archivo Excel debe ser .xlsx"
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

        if (
            not procesar_ovi
            and not procesar_ova
        ):

            return generar_html(
                error=(
                    "Selecciona OVI y/o OVA."
                )
            )

        contenido_excel = await excel.read()

        contenido_zip = await recursos.read()

        hojas = analizar_excel(
            contenido_excel,
            procesar_ovi,
            procesar_ova,
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

        ruta_reporte = os.path.join(
            RESULTADOS_DIR,
            f"resultado_{trabajo_id}.csv",
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
        }

        resultado = {
            "hojas":
                hojas,

            "zip":
                zip_info,
        }

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
):

    trabajo = TRABAJOS.get(
        trabajo_id
    )

    if not trabajo:

        return HTMLResponse(
            """
            <h2>Trabajo no encontrado.</h2>
            <a href="/">Volver</a>
            """,
            status_code=404,
        )

    if trabajo[
        "etapa"
    ] not in [
        "analizado",
        "error",
    ]:

        return HTMLResponse(
            """
            <h2>
                Este trabajo ya fue iniciado.
            </h2>
            """
        )

    trabajo[
        "etapa"
    ] = "iniciando"

    trabajo[
        "mensaje"
    ] = "Iniciando Playwright..."

    trabajo[
        "terminado"
    ] = False

    background_tasks.add_task(
        ejecutar_cargue,
        trabajo["ruta_excel"],
        trabajo["ruta_zip"],
        trabajo["carpeta_temp"],
        trabajo["ruta_reporte"],
        trabajo["procesar_ovi"],
        trabajo["procesar_ova"],
        trabajo,
    )

    return HTMLResponse(
        f"""
        <!DOCTYPE html>

        <html lang="es">

        <head>

            <meta charset="UTF-8">

            <title>
                Cargue PRIZMA
            </title>

            <style>

                body {{
                    font-family: Arial, sans-serif;
                    background: #f4f6fb;
                    margin: 0;
                    color: #1f2937;
                }}

                .caja {{
                    width: 780px;
                    max-width:
                        calc(100% - 40px);
                    margin: 70px auto;
                    background: white;
                    padding: 35px;
                    border-radius: 16px;
                    box-shadow:
                        0 8px 30px
                        rgba(0,0,0,.08);
                }}

                .barra {{
                    height: 20px;
                    background: #e5e7eb;
                    border-radius: 20px;
                    overflow: hidden;
                    margin: 25px 0;
                }}

                .progreso {{
                    height: 100%;
                    width: 0%;
                    background: #16a34a;
                    transition: width .4s;
                }}

                .numeros {{
                    display: grid;
                    grid-template-columns:
                        repeat(3, 1fr);
                    gap: 15px;
                    margin-top: 20px;
                }}

                .numero {{
                    background: #f8fafc;
                    border-radius: 10px;
                    padding: 15px;
                    text-align: center;
                }}

                .numero span {{
                    display: block;
                    color: #6b7280;
                    margin-bottom: 5px;
                }}

                .numero strong {{
                    font-size: 24px;
                }}

                .estado {{
                    padding: 16px;
                    background: #f8fafc;
                    border-radius: 10px;
                    min-height: 55px;
                }}

                .porcentaje {{
                    text-align: right;
                    color: #6b7280;
                }}

                a {{
                    display: inline-block;
                    margin-top: 25px;
                    padding: 13px 20px;
                    background: #111827;
                    color: white;
                    text-decoration: none;
                    border-radius: 8px;
                    font-weight: bold;
                }}

                .final {{
                    margin-top: 25px;
                    padding: 18px;
                    background: #ecfdf5;
                    border-radius: 10px;
                }}

            </style>

        </head>

        <body>

            <div class="caja">

                <h1>
                    Auto Prizma Pro
                </h1>

                <h2>
                    Cargue de curso
                </h2>

                <div
                    id="estado"
                    class="estado"
                >
                    Iniciando...
                </div>

                <div class="barra">

                    <div
                        id="progreso"
                        class="progreso"
                    ></div>

                </div>

                <div
                    id="porcentaje"
                    class="porcentaje"
                >
                    0%
                </div>

                <div class="numeros">

                    <div class="numero">

                        <span>
                            Procesadas
                        </span>

                        <strong id="procesadas">
                            0
                        </strong>

                    </div>

                    <div class="numero">

                        <span>
                            Exitosas
                        </span>

                        <strong id="exitosas">
                            0
                        </strong>

                    </div>

                    <div class="numero">

                        <span>
                            Errores
                        </span>

                        <strong id="errores">
                            0
                        </strong>

                    </div>

                </div>

                <div id="final"></div>

            </div>

            <script>

                const trabajoId =
                    "{trabajo_id}";

                async function revisarEstado() {{

                    try {{

                        const respuesta =
                            await fetch(
                                "/estado/" +
                                trabajoId
                            );

                        const datos =
                            await respuesta.json();

                        document.getElementById(
                            "estado"
                        ).innerText =
                            datos.mensaje;

                        document.getElementById(
                            "procesadas"
                        ).innerText =
                            datos.procesadas;

                        document.getElementById(
                            "exitosas"
                        ).innerText =
                            datos.exitosas;

                        document.getElementById(
                            "errores"
                        ).innerText =
                            datos.errores;

                        let porcentaje = 0;

                        if (datos.total > 0) {{

                            porcentaje =
                                (
                                    datos.procesadas
                                    /
                                    datos.total
                                ) * 100;

                        }}

                        porcentaje =
                            Math.min(
                                100,
                                porcentaje
                            );

                        document.getElementById(
                            "progreso"
                        ).style.width =
                            porcentaje + "%";

                        document.getElementById(
                            "porcentaje"
                        ).innerText =
                            Math.round(
                                porcentaje
                            ) + "%";

                        if (datos.terminado) {{

                            let html =
                                '<div class="final">' +
                                '<h3>Proceso terminado</h3>' +
                                '<p>' +
                                datos.mensaje +
                                '</p>';

                            if (
                                datos.reporte_disponible
                            ) {{

                                html +=
                                    '<a href="/reporte/' +
                                    trabajoId +
                                    '">' +
                                    'Descargar reporte CSV' +
                                    '</a>';

                            }}

                            html += '</div>';

                            document.getElementById(
                                "final"
                            ).innerHTML =
                                html;

                            return;

                        }}

                        setTimeout(
                            revisarEstado,
                            1500
                        );

                    }}
                    catch (error) {{

                        document.getElementById(
                            "estado"
                        ).innerText =
                            "Error consultando estado.";

                        setTimeout(
                            revisarEstado,
                            3000
                        );

                    }}

                }}

                revisarEstado();

            </script>

        </body>

        </html>
        """
    )


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

        "reporte_disponible":
            os.path.isfile(
                trabajo["ruta_reporte"]
            ),
    }


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
        filename="resultado_prizma.csv",
    )