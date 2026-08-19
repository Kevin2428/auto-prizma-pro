from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

from openpyxl import load_workbook

import csv
import os
import re
import zipfile
import unicodedata
import shutil
import traceback


# ============================================================
# CONFIGURACIÓN
# ============================================================

URL_PRIZMA = "https://admin.prizma.site/inicio-sesion"

VERSION_SCRIPT = "PRUEBA_H5P_NUEVO_GESTION_FINANCIERA_V2"

BUILD_INTERNO = "INTERFAZ_WEB_LOGIN_AUTOMATICO_02"


TEXTOS_DESCRIPCION_A_BORRAR = [
    "NO_DISPONIBLE",
    "NO_DISPOINBLE",
]


# ============================================================
# ESTADO DEL TRABAJO
# ============================================================

def actualizar_estado(
    estado,
    etapa=None,
    mensaje=None,
    total=None,
    procesadas=None,
    exitosas=None,
    errores=None,
    terminado=None,
):

    if etapa is not None:
        estado["etapa"] = etapa

    if mensaje is not None:
        estado["mensaje"] = mensaje

    if total is not None:
        estado["total"] = total

    if procesadas is not None:
        estado["procesadas"] = procesadas

    if exitosas is not None:
        estado["exitosas"] = exitosas

    if errores is not None:
        estado["errores"] = errores

    if terminado is not None:
        estado["terminado"] = terminado


# ============================================================
# NORMALIZACIÓN
# ============================================================

def normalizar_texto(texto):

    if texto is None:
        return ""

    texto = str(texto)

    texto = unicodedata.normalize(
        "NFD",
        texto,
    )

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    texto = texto.lower()

    texto = re.sub(
        r"[^a-z0-9]+",
        " ",
        texto,
    )

    return " ".join(
        texto.split()
    )


def normalizar_categoria(categoria):

    valor = normalizar_texto(
        categoria
    )

    if valor == "ovi":
        return "OVI"

    if valor == "ova":
        return "OVA"

    return None


# ============================================================
# TIPO DE ARCHIVO
# ============================================================

def determinar_tipo_archivo(
    categoria,
    tipo_recurso,
    enlace,
):

    categoria_n = normalizar_texto(
        categoria
    )

    tipo_n = normalizar_texto(
        tipo_recurso
    )

    enlace_n = normalizar_texto(
        enlace
    )

    # SOLO OVI / OVA
    if categoria_n not in [
        "ovi",
        "ova",
    ]:
        return None

    # VIDEOS EXCLUIDOS
    if tipo_n in [
        "video intro",
        "video cierre",
        "video a camara",
    ]:
        return None

    # OVA SIEMPRE H5P
    if categoria_n == "ova":
        return "H5P"

    # OVI PDF
    if "pdf" in tipo_n:
        return "PDF"

    if "pdf" in enlace_n:
        return "PDF"

    # OVI H5P
    if "h5p" in tipo_n:
        return "H5P"

    if "h5p" in enlace_n:
        return "H5P"

    # INFOGRAFÍA OVI
    if "infografia" in tipo_n:
        return "H5P"

    return None


# ============================================================
# LEER ACTIVIDADES DEL EXCEL
# ============================================================

def leer_actividades_excel(
    ruta_excel,
    procesar_ovi=True,
    procesar_ova=True,
):

    libro = load_workbook(
        ruta_excel,
        data_only=True,
    )

    actividades = []

    for nombre_hoja in libro.sheetnames:

        hoja = libro[
            nombre_hoja
        ]

        fila_cabecera = None

        # ----------------------------------------------------
        # BUSCAR CABECERA
        # ----------------------------------------------------

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

        programa = (
            str(programa).strip()
            if programa
            else ""
        )

        curso = (
            str(curso).strip()
            if curso
            else nombre_hoja
        )

        # ----------------------------------------------------
        # ACTIVIDADES
        # ----------------------------------------------------

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

            # RETOS / CHALLENGE / OTRAS CATEGORÍAS
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

            actividades.append(
                {
                    "fila_excel": fila,
                    "hoja": nombre_hoja,
                    "programa": programa,
                    "curso": curso,

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

                    "categoria_prizma":
                        categoria_prizma,

                    "tipo_recurso": (
                        str(tipo_recurso).strip()
                        if tipo_recurso
                        else ""
                    ),

                    "tipo_archivo":
                        tipo_archivo,

                    "enlace": (
                        str(enlace).strip()
                        if enlace
                        else ""
                    ),
                }
            )

    return actividades


# ============================================================
# ÍNDICE DE RECURSOS
# ============================================================

def crear_indice_recursos(
    ruta_zip,
):

    indice = []

    with zipfile.ZipFile(
        ruta_zip,
        "r",
    ) as zip_ref:

        for miembro in zip_ref.namelist():

            if miembro.endswith("/"):
                continue

            nombre = os.path.basename(
                miembro
            )

            extension = os.path.splitext(
                nombre
            )[1].lower()

            if extension not in [
                ".h5p",
                ".pdf",
            ]:
                continue

            indice.append(
                {
                    "nombre": nombre,
                    "extension": extension,
                    "zip": ruta_zip,
                    "miembro": miembro,
                }
            )

    return indice


# ============================================================
# PUNTUAR RECURSO
# ============================================================

def puntuar_recurso(
    nombre_archivo,
    actividad,
):

    nombre_sin_extension = os.path.splitext(
        os.path.basename(
            nombre_archivo
        )
    )[0]

    archivo_n = normalizar_texto(
        nombre_sin_extension
    )

    actividad_n = normalizar_texto(
        actividad["nombre"]
    )

    referencia_n = normalizar_texto(
        actividad["enlace"]
    )

    puntuacion = 0

    if archivo_n == actividad_n:
        puntuacion += 300

    if (
        actividad_n
        and actividad_n in archivo_n
    ):
        puntuacion += 180

    if (
        archivo_n
        and archivo_n in referencia_n
    ):
        puntuacion += 80

    if (
        referencia_n
        and referencia_n in archivo_n
    ):
        puntuacion += 120

    palabras_actividad = {
        palabra
        for palabra in actividad_n.split()
        if len(palabra) >= 4
    }

    palabras_archivo = set(
        archivo_n.split()
    )

    if palabras_actividad:

        comunes = (
            palabras_actividad
            & palabras_archivo
        )

        porcentaje = (
            len(comunes)
            / len(palabras_actividad)
        )

        puntuacion += int(
            porcentaje * 100
        )

    return puntuacion


# ============================================================
# RESOLVER RECURSO
# ============================================================

def resolver_recurso(
    actividad,
    indice_recursos,
    carpeta_temp,
):

    print()
    print("--------------------------------------")
    print("BUSCANDO RECURSO")
    print("--------------------------------------")

    print(
        "Actividad:",
        actividad["nombre"],
    )

    print(
        "Tipo esperado:",
        actividad["tipo_archivo"],
    )

    print(
        "Referencia:",
        actividad["enlace"],
    )

    if actividad[
        "tipo_archivo"
    ] == "H5P":

        extension = ".h5p"

    elif actividad[
        "tipo_archivo"
    ] == "PDF":

        extension = ".pdf"

    else:

        return (
            None,
            "ERROR_TIPO_RECURSO_NO_SOPORTADO",
        )

    candidatos = []

    for recurso in indice_recursos:

        if recurso[
            "extension"
        ] != extension:
            continue

        puntuacion = puntuar_recurso(
            recurso["nombre"],
            actividad,
        )

        if puntuacion < 100:
            continue

        candidato = recurso.copy()

        candidato[
            "puntuacion"
        ] = puntuacion

        candidatos.append(
            candidato
        )

    print(
        "Candidatos:",
        len(candidatos),
    )

    if not candidatos:

        return (
            None,
            "ERROR_RECURSO_NO_ENCONTRADO",
        )

    mejor_puntuacion = max(
        candidato["puntuacion"]
        for candidato in candidatos
    )

    mejores = [
        candidato
        for candidato in candidatos
        if candidato["puntuacion"]
        == mejor_puntuacion
    ]

    print(
        "Mejor puntuación:",
        mejor_puntuacion,
    )

    for candidato in mejores:

        print(
            "  →",
            candidato["nombre"],
            "|",
            candidato["puntuacion"],
        )

    # SEGURIDAD: NO ADIVINAR
    if len(mejores) != 1:

        return (
            None,
            "ERROR_RECURSO_DUPLICADO",
        )

    elegido = mejores[0]

    carpeta_fila = os.path.join(
        carpeta_temp,
        "fila_" + str(
            actividad["fila_excel"]
        ),
    )

    os.makedirs(
        carpeta_fila,
        exist_ok=True,
    )

    ruta_destino = os.path.join(
        carpeta_fila,
        elegido["nombre"],
    )

    try:

        with zipfile.ZipFile(
            elegido["zip"],
            "r",
        ) as zip_ref:

            with zip_ref.open(
                elegido["miembro"]
            ) as origen:

                with open(
                    ruta_destino,
                    "wb",
                ) as destino:

                    shutil.copyfileobj(
                        origen,
                        destino,
                    )

    except Exception:

        return (
            None,
            "ERROR_EXTRAYENDO_RECURSO",
        )

    if not os.path.isfile(
        ruta_destino
    ):

        return (
            None,
            "ERROR_EXTRAYENDO_RECURSO",
        )

    print()
    print(
        "✅ Recurso seleccionado:",
        elegido["nombre"],
    )

    return (
        {
            "ruta":
                ruta_destino,

            "nombre_original":
                elegido["nombre"],

            "puntuacion":
                mejor_puntuacion,
        },
        None,
    )


# ============================================================
# REPORTE
# ============================================================

def guardar_resultado(
    ruta_reporte,
    actividad,
    resultado,
    observacion,
    recurso="",
):

    existe = os.path.isfile(
        ruta_reporte
    )

    with open(
        ruta_reporte,
        "a",
        newline="",
        encoding="utf-8-sig",
    ) as archivo:

        escritor = csv.writer(
            archivo
        )

        if not existe:

            escritor.writerow(
                [
                    "Fila Excel",
                    "Programa",
                    "Curso",
                    "Semana",
                    "Unidad",
                    "Actividad",
                    "Categoría",
                    "Tipo",
                    "Recurso",
                    "Resultado",
                    "Observación",
                ]
            )

        escritor.writerow(
            [
                actividad["fila_excel"],
                actividad["programa"],
                actividad["curso"],
                actividad["semana"],
                actividad["unidad"],
                actividad["nombre"],
                actividad["categoria_prizma"],
                actividad["tipo_archivo"],
                recurso,
                resultado,
                observacion,
            ]
        )


# ============================================================
# OVERLAY POST-GUARDADO
# ============================================================

def desbloquear_interfaz_post_guardado(
    pagina,
):

    pagina.wait_for_timeout(
        700
    )

    overlay = pagina.locator(
        "div.MuiBox-root.css-15m6u24"
    )

    try:

        for i in range(
            overlay.count()
        ):

            elemento = overlay.nth(i)

            if not elemento.is_visible():
                continue

            print(
                "✅ Overlay detectado."
            )

            try:

                elemento.click(
                    position={
                        "x": 5,
                        "y": 5,
                    },
                    force=True,
                )

            except Exception:

                pagina.mouse.click(
                    720,
                    450,
                )

            pagina.wait_for_timeout(
                700
            )

            return True

    except Exception:
        pass

    try:

        pagina.mouse.click(
            720,
            450,
        )

        pagina.wait_for_timeout(
            400
        )

    except Exception:
        pass

    return True


# ============================================================
# CANCELAR EDICIÓN
# ============================================================

def cancelar_edicion_segura(
    pagina,
):

    try:

        boton = pagina.get_by_role(
            "button",
            name="Cancelar",
            exact=True,
        )

        if boton.count() == 0:
            return False

        if not boton.first.is_visible():
            return False

        boton.first.click(
            timeout=10000
        )

        pagina.wait_for_timeout(
            900
        )

        desbloquear_interfaz_post_guardado(
            pagina
        )

        return True

    except Exception:

        return False


# ============================================================
# ASEGURAR LISTADO
# ============================================================

def asegurar_listado(
    pagina,
):

    desbloquear_interfaz_post_guardado(
        pagina
    )

    buscador = pagina.locator(
        'input[placeholder="Buscar..."]'
    )

    try:

        if (
            buscador.count() > 0
            and buscador.first.is_visible()
        ):
            return True

    except Exception:
        pass

    try:

        cancelar = pagina.get_by_role(
            "button",
            name="Cancelar",
            exact=True,
        )

        if (
            cancelar.count() > 0
            and cancelar.first.is_visible()
        ):

            cancelar.first.click(
                timeout=10000
            )

            pagina.wait_for_timeout(
                1000
            )

            desbloquear_interfaz_post_guardado(
                pagina
            )

    except Exception:
        pass

    try:

        buscador.wait_for(
            state="visible",
            timeout=10000,
        )

        return True

    except Exception:

        return False


# ============================================================
# ENTRAR A CATEGORÍA
# ============================================================

def entrar_categoria(
    pagina,
    categoria,
):

    if categoria not in [
        "OVI",
        "OVA",
    ]:

        return False

    try:

        desbloquear_interfaz_post_guardado(
            pagina
        )

        pestana = pagina.locator(
            "label.tab",
            has_text=categoria,
        ).first

        pestana.wait_for(
            state="visible",
            timeout=30000,
        )

        pestana.click(
            timeout=10000
        )

        pagina.wait_for_timeout(
            1200
        )

        return True

    except Exception:

        return False


# ============================================================
# ENCONTRAR FILA
# ============================================================

def encontrar_fila_desde_elemento(
    elemento,
):

    for nivel in range(
        1,
        10,
    ):

        try:

            padre = elemento.locator(
                "xpath=" + "/.." * nivel
            )

            if padre.locator(
                "button"
            ).count() == 3:

                return padre

        except Exception:
            pass

    return None


def obtener_id_temporal_fila(
    fila,
):

    try:

        return fila.evaluate(
            """
            (element) => {

                if (!element.dataset.prizmaScanId) {

                    element.dataset.prizmaScanId =
                        "scan-" +
                        Math.random()
                        .toString(36)
                        .slice(2);
                }

                return element.dataset.prizmaScanId;
            }
            """
        )

    except Exception:

        return None


def obtener_nombre_fila(
    texto,
):

    lineas = [
        linea.strip()
        for linea in texto.splitlines()
        if linea.strip()
    ]

    if len(lineas) < 2:
        return ""

    return lineas[1]


# ============================================================
# ANALIZAR RESULTADOS
# ============================================================

def analizar_resultados_pagina(
    pagina,
    actividad,
):

    elementos = pagina.get_by_text(
        actividad["nombre"],
        exact=False,
    )

    filas = {}

    for i in range(
        elementos.count()
    ):

        try:

            elemento = elementos.nth(i)

            if not elemento.is_visible():
                continue

            fila = encontrar_fila_desde_elemento(
                elemento
            )

            if fila is None:
                continue

            identificador = obtener_id_temporal_fila(
                fila
            )

            if identificador:

                filas[
                    identificador
                ] = fila

        except Exception:
            pass

    nombre_objetivo = normalizar_texto(
        actividad["nombre"]
    )

    semana_objetivo = normalizar_texto(
        actividad["semana"]
    )

    unidad_objetivo = normalizar_texto(
        actividad["unidad"]
    )

    programa_objetivo = normalizar_texto(
        actividad["programa"]
    )

    resultados = []

    for fila in filas.values():

        try:

            texto = fila.inner_text().strip()

        except Exception:
            continue

        texto_n = normalizar_texto(
            texto
        )

        nombre_fila = normalizar_texto(
            obtener_nombre_fila(
                texto
            )
        )

        cumple_nombre = (
            nombre_fila
            == nombre_objetivo
        )

        cumple_semana = (
            semana_objetivo
            in texto_n
        )

        cumple_unidad = (
            unidad_objetivo
            in texto_n
        )

        cumple_programa = (
            programa_objetivo
            in texto_n
        )

        if actividad[
            "categoria_prizma"
        ] == "OVI":

            cumple_categoria = (
                "ovi" in texto_n
            )

        elif actividad[
            "categoria_prizma"
        ] == "OVA":

            cumple_categoria = (
                "ova" in texto_n
            )

        else:

            cumple_categoria = False

        coincide = (
            cumple_nombre
            and cumple_semana
            and cumple_unidad
            and cumple_programa
            and cumple_categoria
        )

        resultados.append(
            {
                "fila": fila,
                "texto": texto,
                "coincide": coincide,
            }
        )

    return resultados


def obtener_firma_pagina(
    pagina,
    actividad,
):

    resultados = analizar_resultados_pagina(
        pagina,
        actividad,
    )

    return "||".join(
        sorted(
            normalizar_texto(
                resultado["texto"]
            )
            for resultado in resultados
        )
    )


# ============================================================
# PAGINACIÓN
# ============================================================

def obtener_paginas_numericas(
    pagina,
):

    numeros = set()

    botones = pagina.locator(
        "button"
    )

    for i in range(
        botones.count()
    ):

        try:

            boton = botones.nth(i)

            if not boton.is_visible():
                continue

            texto = boton.inner_text().strip()

            if texto.isdigit():

                numero = int(
                    texto
                )

                if 1 <= numero <= 200:

                    numeros.add(
                        numero
                    )

            aria = (
                boton.get_attribute(
                    "aria-label"
                )
                or ""
            )

            coincidencia = re.search(
                r"(?:page|página|pagina)\s*(\d+)",
                aria,
                flags=re.IGNORECASE,
            )

            if coincidencia:

                numero = int(
                    coincidencia.group(1)
                )

                if 1 <= numero <= 200:

                    numeros.add(
                        numero
                    )

        except Exception:
            pass

    return sorted(
        numeros
    )


def ir_a_pagina_numero(
    pagina,
    numero,
):

    botones = pagina.locator(
        "button"
    )

    # POR TEXTO
    for i in range(
        botones.count()
    ):

        try:

            boton = botones.nth(i)

            if not boton.is_visible():
                continue

            texto = boton.inner_text().strip()

            if texto != str(
                numero
            ):
                continue

            boton.click(
                timeout=5000
            )

            pagina.wait_for_timeout(
                1200
            )

            return True

        except Exception:
            pass

    # POR ARIA
    for i in range(
        botones.count()
    ):

        try:

            boton = botones.nth(i)

            if not boton.is_visible():
                continue

            aria = normalizar_texto(
                boton.get_attribute(
                    "aria-label"
                )
            )

            if (
                f"page {numero}" not in aria
                and
                f"pagina {numero}" not in aria
            ):
                continue

            boton.click(
                timeout=5000
            )

            pagina.wait_for_timeout(
                1200
            )

            return True

        except Exception:
            pass

    return False


def encontrar_boton_siguiente(
    pagina,
):

    botones = pagina.locator(
        "button"
    )

    for i in range(
        botones.count()
    ):

        try:

            boton = botones.nth(i)

            if not boton.is_visible():
                continue

            if not boton.is_enabled():
                continue

            texto = normalizar_texto(
                boton.inner_text()
            )

            aria = normalizar_texto(
                boton.get_attribute(
                    "aria-label"
                )
            )

            title = normalizar_texto(
                boton.get_attribute(
                    "title"
                )
            )

            combinado = (
                texto
                + " "
                + aria
                + " "
                + title
            )

            if any(
                valor in combinado
                for valor in [
                    "next page",
                    "pagina siguiente",
                    "siguiente",
                    "go to next page",
                ]
            ):

                return boton

        except Exception:
            pass

    return None


def asegurar_pagina_1(
    pagina,
):

    paginas = obtener_paginas_numericas(
        pagina
    )

    if 1 not in paginas:
        return True

    return ir_a_pagina_numero(
        pagina,
        1,
    )


# ============================================================
# BUSCAR ACTIVIDAD
# ============================================================

def buscar_actividad_correcta(
    pagina,
    actividad,
):

    print()
    print(
        "Buscando actividad:",
        actividad["nombre"],
    )

    print(
        "Programa:",
        actividad["programa"],
    )

    print(
        "Semana:",
        actividad["semana"],
    )

    print(
        "Unidad:",
        actividad["unidad"],
    )

    print(
        "Categoría:",
        actividad["categoria_prizma"],
    )

    buscador = pagina.locator(
        'input[placeholder="Buscar..."]'
    )

    buscador.wait_for(
        state="visible",
        timeout=30000,
    )

    buscador.fill("")

    pagina.wait_for_timeout(
        500
    )

    asegurar_pagina_1(
        pagina
    )

    buscador.fill(
        actividad["nombre"]
    )

    pagina.wait_for_timeout(
        1800
    )

    asegurar_pagina_1(
        pagina
    )

    pagina.wait_for_timeout(
        700
    )

    coincidencias = []

    firmas_visitadas = set()

    numero_logico = 1

    while numero_logico <= 50:

        firma_actual = obtener_firma_pagina(
            pagina,
            actividad,
        )

        if firma_actual in firmas_visitadas:
            break

        firmas_visitadas.add(
            firma_actual
        )

        resultados = analizar_resultados_pagina(
            pagina,
            actividad,
        )

        print(
            "Página",
            numero_logico,
            "- candidatos:",
            len(resultados),
        )

        for resultado in resultados:

            if resultado[
                "coincide"
            ]:

                coincidencias.append(
                    {
                        "pagina":
                            numero_logico,

                        "texto":
                            resultado["texto"],
                    }
                )

        siguiente_numero = (
            numero_logico + 1
        )

        paginas = obtener_paginas_numericas(
            pagina
        )

        if siguiente_numero in paginas:

            if ir_a_pagina_numero(
                pagina,
                siguiente_numero,
            ):

                firma_nueva = obtener_firma_pagina(
                    pagina,
                    actividad,
                )

                if firma_nueva != firma_actual:

                    numero_logico += 1
                    continue

        siguiente = encontrar_boton_siguiente(
            pagina
        )

        if siguiente is None:
            break

        try:

            siguiente.click(
                timeout=5000
            )

            pagina.wait_for_timeout(
                1200
            )

        except Exception:
            break

        firma_nueva = obtener_firma_pagina(
            pagina,
            actividad,
        )

        if firma_nueva == firma_actual:
            break

        numero_logico += 1

    print(
        "Coincidencias exactas:",
        len(coincidencias),
    )

    if len(coincidencias) == 0:

        return (
            None,
            "ERROR_ACTIVIDAD_NO_ENCONTRADA",
        )

    if len(coincidencias) > 1:

        return (
            None,
            "ERROR_ACTIVIDAD_DUPLICADA",
        )

    pagina_objetivo = coincidencias[
        0
    ][
        "pagina"
    ]

    # --------------------------------------------------------
    # RECUPERAR PÁGINA
    # --------------------------------------------------------

    if not ir_a_pagina_numero(
        pagina,
        pagina_objetivo,
    ):

        asegurar_pagina_1(
            pagina
        )

        pagina_actual = 1

        while pagina_actual < pagina_objetivo:

            siguiente_numero = (
                pagina_actual + 1
            )

            if ir_a_pagina_numero(
                pagina,
                siguiente_numero,
            ):

                pagina_actual += 1
                continue

            siguiente = encontrar_boton_siguiente(
                pagina
            )

            if siguiente is None:

                return (
                    None,
                    "ERROR_RECUPERANDO_PAGINA",
                )

            try:

                siguiente.click(
                    timeout=5000
                )

                pagina.wait_for_timeout(
                    1200
                )

                pagina_actual += 1

            except Exception:

                return (
                    None,
                    "ERROR_RECUPERANDO_PAGINA",
                )

    pagina.wait_for_timeout(
        700
    )

    resultados_finales = analizar_resultados_pagina(
        pagina,
        actividad,
    )

    finales = [
        resultado
        for resultado in resultados_finales
        if resultado[
            "coincide"
        ]
    ]

    if len(finales) == 0:

        return (
            None,
            "ERROR_RECUPERANDO_FILA",
        )

    if len(finales) > 1:

        return (
            None,
            "ERROR_ACTIVIDAD_DUPLICADA",
        )

    print(
        "✅ Fila recuperada."
    )

    return (
        finales[0]["fila"],
        None,
    )


# ============================================================
# ABRIR EDICIÓN
# ============================================================

def abrir_edicion(
    pagina,
    fila,
):

    botones = fila.locator(
        "button"
    )

    if botones.count() != 3:

        print(
            "❌ Cantidad inesperada de botones:",
            botones.count(),
        )

        return False

    try:

        botones.nth(1).click(
            timeout=10000
        )

    except PlaywrightTimeoutError:

        desbloquear_interfaz_post_guardado(
            pagina
        )

        try:

            botones.nth(1).click(
                timeout=10000
            )

        except Exception:
            return False

    except Exception:
        return False

    pagina.wait_for_timeout(
        1800
    )

    return True


# ============================================================
# CAMPO PRINCIPAL "RECURSO"
# ============================================================
#
# TODOS LOS ARCHIVOS SE CARGAN EN RECURSO.
#
# H5P -> RECURSO
# PDF -> RECURSO
#
# MATERIAL DESCARGABLE NO SE TOCA.
# ============================================================

def detectar_campo_recurso(
    pagina,
):

    archivos = pagina.locator(
        'input[type="file"]'
    )

    candidatos_recurso = []

    print()
    print(
        "Campos file detectados:",
        archivos.count(),
    )

    for i in range(
        archivos.count()
    ):

        archivo = archivos.nth(i)

        accept = (
            archivo.get_attribute(
                "accept"
            )
            or ""
        )

        accept_n = accept.lower()

        print()
        print(
            f"Campo #{i + 1}:"
        )

        print(
            "accept =",
            accept,
        )

        # El campo Recurso admite .h5p.
        if ".h5p" in accept_n:

            candidatos_recurso.append(
                archivo
            )

            print(
                "→ RECURSO PRINCIPAL"
            )

            continue

        # Material descargable.
        if accept_n.strip() == ".pdf":

            print(
                "→ MATERIAL DESCARGABLE - IGNORADO"
            )

            continue

        print(
            "→ CAMPO NO UTILIZADO"
        )

    if len(
        candidatos_recurso
    ) == 0:

        print(
            "❌ No se encontró RECURSO."
        )

        return None

    if len(
        candidatos_recurso
    ) > 1:

        print(
            "❌ Varios candidatos para RECURSO."
        )

        return None

    print()
    print(
        "✅ Campo RECURSO identificado."
    )

    return candidatos_recurso[0]


# ============================================================
# DESCRIPCIÓN
# ============================================================

def limpiar_descripcion(
    pagina,
):

    editores = pagina.locator(
        '[contenteditable="true"]'
    )

    objetivos = [
        normalizar_texto(
            texto
        )
        for texto in TEXTOS_DESCRIPCION_A_BORRAR
    ]

    exactos = []
    peligrosos = []

    for i in range(
        editores.count()
    ):

        try:

            editor = editores.nth(i)

            if not editor.is_visible():
                continue

            texto_original = (
                editor.inner_text().strip()
            )

            texto_n = normalizar_texto(
                texto_original
            )

            if texto_n in objetivos:

                exactos.append(
                    editor
                )

                continue

            for objetivo in objetivos:

                if (
                    objetivo
                    and objetivo in texto_n
                ):

                    peligrosos.append(
                        texto_original
                    )

        except Exception:
            pass

    if peligrosos:

        return (
            False,
            "ERROR_DESCRIPCION_CONTENIDO_ADICIONAL",
        )

    if len(exactos) == 0:

        return (
            True,
            None,
        )

    if len(exactos) > 1:

        return (
            False,
            "ERROR_DESCRIPCION_MULTIPLE",
        )

    editor = exactos[0]

    try:

        editor.click()

        editor.press(
            "Control+A"
        )

        editor.press(
            "Backspace"
        )

        pagina.wait_for_timeout(
            500
        )

        if normalizar_texto(
            editor.inner_text()
        ):

            return (
                False,
                "ERROR_LIMPIANDO_DESCRIPCION",
            )

        print(
            "✅ NO_DISPONIBLE eliminado."
        )

        return (
            True,
            None,
        )

    except Exception:

        return (
            False,
            "ERROR_LIMPIANDO_DESCRIPCION",
        )


# ============================================================
# BOTÓN GUARDAR
# ============================================================

def encontrar_boton_guardar(
    pagina,
):

    botones = pagina.get_by_role(
        "button",
        name="Editar",
        exact=True,
    )

    for i in range(
        botones.count()
    ):

        try:

            boton = botones.nth(i)

            if not boton.is_visible():
                continue

            if not boton.is_enabled():
                continue

            tipo = (
                boton.get_attribute(
                    "type"
                )
                or ""
            ).lower()

            if tipo == "submit":

                return boton

        except Exception:
            pass

    return None


# ============================================================
# ARCHIVO VISIBLE
# ============================================================

def archivo_visible(
    pagina,
    nombre_archivo,
):

    referencia = pagina.get_by_text(
        nombre_archivo,
        exact=False,
    )

    for i in range(
        referencia.count()
    ):

        try:

            if referencia.nth(i).is_visible():

                return True

        except Exception:
            pass

    return False


# ============================================================
# PATCH
# ============================================================

def esperar_patch(
    pagina,
    respuestas,
    timeout_ms=15000,
):

    esperado = 0

    while esperado < timeout_ms:

        for respuesta in respuestas:

            if respuesta[
                "metodo"
            ] != "PATCH":
                continue

            if "/academic/activity/" not in respuesta[
                "url"
            ]:
                continue

            return respuesta

        pagina.wait_for_timeout(
            250
        )

        esperado += 250

    return None


# ============================================================
# LOGIN AUTOMÁTICO
# ============================================================

def iniciar_sesion_prizma(
    pagina,
    usuario_prizma,
    contrasena_prizma,
):

    print()
    print("--------------------------------------")
    print("LOGIN AUTOMÁTICO PRIZMA")
    print("--------------------------------------")

    pagina.goto(
        URL_PRIZMA,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    campo_usuario = pagina.locator(
        'input[name="identification_number"]'
    )

    campo_contrasena = pagina.locator(
        'input[name="password"]'
    )

    boton_login = pagina.get_by_role(
        "button",
        name="Iniciar sesión",
        exact=True,
    )

    campo_usuario.wait_for(
        state="visible",
        timeout=30000,
    )

    campo_contrasena.wait_for(
        state="visible",
        timeout=30000,
    )

    boton_login.wait_for(
        state="visible",
        timeout=30000,
    )

    print(
        "✅ Formulario de login detectado."
    )

    campo_usuario.fill(
        usuario_prizma
    )

    campo_contrasena.fill(
        contrasena_prizma
    )

    print(
        "✅ Credenciales introducidas."
    )

    print(
        "Iniciando sesión..."
    )

    boton_login.click(
        timeout=10000
    )

    # --------------------------------------------------------
    # CONFIRMAR PANEL
    # --------------------------------------------------------

    bienvenida = pagina.get_by_text(
        "Bienvenido a Prizma admin",
        exact=False,
    )

    try:

        bienvenida.wait_for(
            state="visible",
            timeout=60000,
        )

    except Exception:

        try:

            texto_visible = pagina.locator(
                "body"
            ).inner_text()

            print()
            print(
                "No se pudo confirmar el panel."
            )

            print(
                texto_visible[:1200]
            )

        except Exception:
            pass

        raise RuntimeError(
            "ERROR_LOGIN_PRIZMA"
        )

    print(
        "✅ Login automático confirmado."
    )

    print(
        "URL después del login:",
        pagina.url,
    )

    # --------------------------------------------------------
    # ENTRAR A ACTIVIDADES
    # --------------------------------------------------------

    tarjeta_actividades = pagina.get_by_role(
        "button"
    ).filter(
        has_text="Actividades"
    ).first

    tarjeta_actividades.wait_for(
        state="visible",
        timeout=30000,
    )

    print(
        "✅ Tarjeta Actividades detectada."
    )

    tarjeta_actividades.click(
        timeout=10000
    )

    # --------------------------------------------------------
    # CONFIRMAR LISTADO
    # --------------------------------------------------------

    buscador = pagina.locator(
        'input[placeholder="Buscar..."]'
    )

    buscador.wait_for(
        state="visible",
        timeout=60000,
    )

    print(
        "✅ Módulo Actividades abierto."
    )

    return True


# ============================================================
# PROCESAR ACTIVIDAD
# ============================================================

def procesar_actividad(
    pagina,
    actividad,
    indice_recursos,
    carpeta_temp,
    respuestas,
    captura,
):

    print()
    print()
    print("======================================")

    print(
        "ACTIVIDAD:",
        actividad["nombre"],
    )

    print(
        "FILA:",
        actividad["fila_excel"],
    )

    print(
        "CATEGORÍA:",
        actividad["categoria_prizma"],
    )

    print(
        "TIPO:",
        actividad["tipo_archivo"],
    )

    print("======================================")

    # 1. RECURSO

    recurso, error = resolver_recurso(
        actividad,
        indice_recursos,
        carpeta_temp,
    )

    if error:

        return {
            "ok": False,
            "error": error,
            "recurso": "",
        }

    ruta_recurso = recurso[
        "ruta"
    ]

    nombre_recurso = recurso[
        "nombre_original"
    ]

    # 2. LISTADO

    if not asegurar_listado(
        pagina
    ):

        return {
            "ok": False,
            "error":
                "ERROR_NO_SE_PUDO_VOLVER_AL_LISTADO",
            "recurso":
                nombre_recurso,
        }

    # 3. CATEGORÍA

    if not entrar_categoria(
        pagina,
        actividad["categoria_prizma"],
    ):

        return {
            "ok": False,
            "error":
                "ERROR_ENTRANDO_CATEGORIA",
            "recurso":
                nombre_recurso,
        }

    # 4. BUSCAR

    fila, error = buscar_actividad_correcta(
        pagina,
        actividad,
    )

    if error:

        return {
            "ok": False,
            "error": error,
            "recurso":
                nombre_recurso,
        }

    # 5. EDITAR

    if not abrir_edicion(
        pagina,
        fila,
    ):

        return {
            "ok": False,
            "error":
                "ERROR_ABRIENDO_EDICION",
            "recurso":
                nombre_recurso,
        }

    # 6. CAMPO RECURSO

    campo_recurso = detectar_campo_recurso(
        pagina
    )

    if campo_recurso is None:

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_CAMPO_RECURSO_NO_ENCONTRADO",
            "recurso":
                nombre_recurso,
        }

    print()
    print(
        "✅ El archivo será cargado en RECURSO."
    )

    print(
        "Material descargable: IGNORADO"
    )

    # 7. DESCRIPCIÓN

    descripcion_ok, error_descripcion = limpiar_descripcion(
        pagina
    )

    if not descripcion_ok:

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                error_descripcion,
            "recurso":
                nombre_recurso,
        }

    # 8. CARGAR

    print()
    print(
        "Cargando:",
        nombre_recurso,
    )

    try:

        campo_recurso.set_input_files(
            ruta_recurso
        )

    except Exception:

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_SET_INPUT_FILES",
            "recurso":
                nombre_recurso,
        }

    pagina.wait_for_timeout(
        2200
    )

    # 9. VISIBLE

    if not archivo_visible(
        pagina,
        nombre_recurso,
    ):

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_RECURSO_NO_VISIBLE",
            "recurso":
                nombre_recurso,
        }

    print(
        "✅ Recurso visible."
    )

    # 10. GUARDAR

    boton_guardar = encontrar_boton_guardar(
        pagina
    )

    if boton_guardar is None:

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_BOTON_GUARDAR",
            "recurso":
                nombre_recurso,
        }

    respuestas.clear()

    captura[
        "activa"
    ] = True

    print(
        "💾 Guardando..."
    )

    try:

        boton_guardar.click(
            timeout=10000
        )

    except Exception:

        captura[
            "activa"
        ] = False

        cancelar_edicion_segura(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_CLIC_GUARDAR",
            "recurso":
                nombre_recurso,
        }

    # 11. PATCH

    respuesta_patch = esperar_patch(
        pagina,
        respuestas,
        timeout_ms=15000,
    )

    captura[
        "activa"
    ] = False

    if respuesta_patch is None:

        asegurar_listado(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_PATCH_NO_CONFIRMADO",
            "recurso":
                nombre_recurso,
        }

    if not (
        200
        <= respuesta_patch[
            "status"
        ]
        < 300
    ):

        asegurar_listado(
            pagina
        )

        return {
            "ok": False,
            "error":
                "ERROR_PATCH_HTTP_"
                + str(
                    respuesta_patch["status"]
                ),
            "recurso":
                nombre_recurso,
        }

    print(
        "✅ PATCH confirmado:",
        respuesta_patch["status"],
    )

    # ========================================================
    # NO REABRIR.
    # DIRECTO A SIGUIENTE ACTIVIDAD.
    # ========================================================

    if not asegurar_listado(
        pagina
    ):

        return {
            "ok": False,
            "error":
                "ERROR_POST_GUARDADO_LISTADO",
            "recurso":
                nombre_recurso,
        }

    return {
        "ok": True,
        "error": "",
        "recurso":
            nombre_recurso,
    }


# ============================================================
# MOTOR COMPLETO
# ============================================================

def ejecutar_cargue(
    ruta_excel,
    ruta_zip,
    carpeta_temp,
    ruta_reporte,
    procesar_ovi,
    procesar_ova,
    usuario_prizma,
    contrasena_prizma,
    estado,
):

    try:

        print()
        print("======================================")
        print("AUTO PRIZMA PRO")
        print("======================================")

        print(
            "VERSION_SCRIPT =",
            VERSION_SCRIPT,
        )

        print(
            "BUILD_INTERNO =",
            BUILD_INTERNO,
        )

        print(
            "URL =",
            URL_PRIZMA,
        )

        print(
            "MODO = CURSO COMPLETO"
        )

        print(
            "NAVEGADOR = HEADLESS"
        )

        print(
            "LOGIN = AUTOMÁTICO"
        )

        print()
        print(
            "H5P -> RECURSO"
        )

        print(
            "PDF -> RECURSO"
        )

        print(
            "MATERIAL DESCARGABLE -> NO TOCAR"
        )

        # ----------------------------------------------------
        # PREPARAR
        # ----------------------------------------------------

        actualizar_estado(
            estado,
            etapa="preparando",
            mensaje=(
                "Leyendo Excel y preparando recursos..."
            ),
            terminado=False,
        )

        actividades = leer_actividades_excel(
            ruta_excel,
            procesar_ovi,
            procesar_ova,
        )

        if not actividades:

            actualizar_estado(
                estado,
                etapa="error",
                mensaje=(
                    "No se encontraron actividades "
                    "OVI/OVA compatibles."
                ),
                terminado=True,
            )

            return

        indice_recursos = crear_indice_recursos(
            ruta_zip
        )

        if not indice_recursos:

            actualizar_estado(
                estado,
                etapa="error",
                mensaje=(
                    "No se encontraron archivos "
                    "H5P/PDF dentro del ZIP."
                ),
                terminado=True,
            )

            return

        actividades_a_procesar = actividades

        print()
        print(
            "Actividades que se procesarán:",
            len(actividades_a_procesar),
        )

        actualizar_estado(
            estado,
            etapa="login",
            mensaje=(
                "Iniciando sesión automáticamente en PRIZMA..."
            ),
            total=len(
                actividades_a_procesar
            ),
            procesadas=0,
            exitosas=0,
            errores=0,
        )

        # ----------------------------------------------------
        # PLAYWRIGHT
        # ----------------------------------------------------

        with sync_playwright() as p:

            navegador = p.chromium.launch(
                headless=True
            )

            pagina = navegador.new_page(
                viewport={
                    "width": 1440,
                    "height": 900,
                }
            )

            respuestas = []

            captura = {
                "activa": False
            }

            def registrar_respuesta(
                response,
            ):

                if not captura[
                    "activa"
                ]:
                    return

                try:

                    metodo = (
                        response.request.method.upper()
                    )

                    if metodo == "GET":
                        return

                    respuestas.append(
                        {
                            "metodo":
                                metodo,

                            "status":
                                response.status,

                            "url":
                                response.url,
                        }
                    )

                except Exception:
                    pass

            pagina.on(
                "response",
                registrar_respuesta,
            )

            # ------------------------------------------------
            # LOGIN AUTOMÁTICO
            # ------------------------------------------------

            iniciar_sesion_prizma(
                pagina,
                usuario_prizma,
                contrasena_prizma,
            )

            actualizar_estado(
                estado,
                etapa="procesando",
                mensaje=(
                    "Login correcto. "
                    "Módulo Actividades abierto."
                ),
            )

            exitosas = 0
            errores = 0
            procesadas = 0

            # ------------------------------------------------
            # TODAS LAS ACTIVIDADES
            # ------------------------------------------------

            for numero, actividad in enumerate(
                actividades_a_procesar,
                start=1,
            ):

                print()
                print()
                print("######################################")

                print(
                    f"ACTIVIDAD {numero} "
                    f"DE {len(actividades_a_procesar)}"
                )

                print("######################################")

                actualizar_estado(
                    estado,
                    etapa="procesando",
                    mensaje=(
                        f"Procesando {numero} de "
                        f"{len(actividades_a_procesar)}: "
                        + actividad["nombre"]
                    ),
                )

                try:

                    resultado = procesar_actividad(
                        pagina,
                        actividad,
                        indice_recursos,
                        carpeta_temp,
                        respuestas,
                        captura,
                    )

                except Exception as e:

                    print()
                    print(
                        traceback.format_exc()
                    )

                    resultado = {
                        "ok": False,

                        "error":
                            "ERROR_NO_CONTROLADO: "
                            + str(e),

                        "recurso": "",
                    }

                    try:

                        cancelar_edicion_segura(
                            pagina
                        )

                    except Exception:
                        pass

                    try:

                        asegurar_listado(
                            pagina
                        )

                    except Exception:
                        pass

                procesadas += 1

                # --------------------------------------------
                # OK
                # --------------------------------------------

                if resultado[
                    "ok"
                ]:

                    exitosas += 1

                    print()
                    print(
                        "✅ ACTIVIDAD COMPLETADA."
                    )

                    guardar_resultado(
                        ruta_reporte,
                        actividad,
                        "OK",
                        (
                            "Carga guardada - "
                            "PATCH confirmado"
                        ),
                        resultado["recurso"],
                    )

                # --------------------------------------------
                # ERROR
                # --------------------------------------------

                else:

                    errores += 1

                    print()
                    print(
                        "❌ ACTIVIDAD CON ERROR:"
                    )

                    print(
                        resultado["error"]
                    )

                    guardar_resultado(
                        ruta_reporte,
                        actividad,
                        "ERROR",
                        resultado["error"],
                        resultado["recurso"],
                    )

                # --------------------------------------------
                # ACTUALIZAR WEB
                # --------------------------------------------

                actualizar_estado(
                    estado,
                    procesadas=procesadas,
                    exitosas=exitosas,
                    errores=errores,
                    mensaje=(
                        f"Procesadas {procesadas} de "
                        f"{len(actividades_a_procesar)}"
                    ),
                )

                try:

                    asegurar_listado(
                        pagina
                    )

                except Exception:
                    pass

                pagina.wait_for_timeout(
                    700
                )

            navegador.close()

        # ----------------------------------------------------
        # FINAL
        # ----------------------------------------------------

        actualizar_estado(
            estado,
            etapa="finalizado",
            mensaje=(
                "Cargue completo finalizado. "
                f"Exitosas: {exitosas}. "
                f"Errores: {errores}."
            ),
            terminado=True,
        )

    except Exception as e:

        print()
        print(
            traceback.format_exc()
        )

        mensaje_error = str(e)

        if "ERROR_LOGIN_PRIZMA" in mensaje_error:

            mensaje_error = (
                "No fue posible iniciar sesión en PRIZMA. "
                "Verifica el usuario y la contraseña."
            )

        actualizar_estado(
            estado,
            etapa="error",
            mensaje=mensaje_error,
            terminado=True,
        )