"""Hace robusta la recuperacion de fila en PRIZMA.

Problema que corrige:
    Tras confirmar la coincidencia exacta, el motor vuelve a filtrar para
    dejar la fila en pantalla. Ese segundo intento esperaba MENOS tiempo
    que la busqueda original (1800/3200 ms contra 2600/4200 ms), se hacia
    UNA sola vez y aceptaba como "resultado" la pantalla anterior si esta
    todavia no se habia refrescado. Si PRIZMA tardaba un poco mas de lo
    normal, el escaneo veia la pagina vacia y devolvia
    ERROR_RECUPERANDO_FILA aunque la actividad si existiera.

Que hace este parche:
    1. La espera de resultados descarta la pantalla anterior (evita leer
       datos viejos) y exige dos lecturas estables.
    2. La recuperacion se reintenta hasta 3 veces con esperas crecientes
       y probando todos los terminos de busqueda que sirvieron.
    3. Un fallo al pasar de pagina ya no aborta: se reintenta.

Uso:
    python parche_recuperar_fila.py
"""

import io
import os
import shutil
import sys

RUTA = "motor_prizma.py"


# ------------------------------------------------------------------
# 1) Espera de resultados: no aceptar la pantalla anterior
# ------------------------------------------------------------------
VIEJO_ESPERA = '''def _esperar_resultados_busqueda(
    pagina,
    actividad,
    timeout_ms=2600,
    intervalo_ms=100,
):
    """Espera solo hasta que aparezcan resultados y la firma quede estable.

    Si no aparece ningun resultado conserva el timeout historico completo,
    evitando falsos negativos por una respuesta lenta de PRIZMA.
    """
    transcurrido = 0
    firma_anterior = None
    estables = 0

    while True:
        firma = obtener_firma_pagina(pagina, actividad)

        if firma:
            if firma == firma_anterior:
                estables += 1
            else:
                estables = 0

            if estables >= 1:
                return firma

        firma_anterior = firma
'''

NUEVO_ESPERA = '''def _esperar_resultados_busqueda(
    pagina,
    actividad,
    timeout_ms=2600,
    intervalo_ms=100,
    firma_previa=None,
):
    """Espera hasta que aparezcan resultados y la firma quede estable.

    firma_previa es la pantalla que habia ANTES de aplicar el filtro.
    Mientras la firma siga siendo esa, el filtro todavia no se aplico y
    seguir adelante significaria leer datos viejos. Ese era el origen de
    ERROR_RECUPERANDO_FILA cuando PRIZMA respondia lento.

    Si no aparece ningun resultado conserva el timeout completo, evitando
    falsos negativos por una respuesta lenta de PRIZMA.
    """
    transcurrido = 0
    firma_anterior = None
    estables = 0

    while True:
        firma = obtener_firma_pagina(pagina, actividad)

        if firma and firma != firma_previa:
            if firma == firma_anterior:
                estables += 1
            else:
                estables = 0

            # Dos lecturas iguales seguidas: la tabla ya dejo de moverse.
            if estables >= 2:
                return firma

        firma_anterior = firma
'''


# ------------------------------------------------------------------
# 2) Recuperacion de fila con reintentos
# ------------------------------------------------------------------
VIEJO_RECUPERACION = '''    buscador.fill("")
    pagina.wait_for_timeout(500)
    asegurar_pagina_1(pagina)
    buscador.fill(termino_encontrado)

    espera_recuperacion = (
        3200
        if actividad["categoria_prizma"] == "CHALLENGE"
        else 1800
    )
    _esperar_resultados_busqueda(
        pagina,
        actividad,
        timeout_ms=espera_recuperacion,
    )

    asegurar_pagina_1(pagina)

    firmas_recuperacion = set()
    numero_recuperacion = 1

    while numero_recuperacion <= 50:

        firma_actual = obtener_firma_pagina(
            pagina,
            actividad,
        )

        if firma_actual in firmas_recuperacion:
            break

        firmas_recuperacion.add(
            firma_actual
        )

        resultados_finales = analizar_resultados_pagina(
            pagina,
            actividad,
        )

        finales = [
            resultado
            for resultado in resultados_finales
            if resultado["coincide"]
        ]

        if len(finales) > 1:
            return (
                None,
                "ERROR_ACTIVIDAD_DUPLICADA",
            )

        if len(finales) == 1:
            print(
                "✅ Fila recuperada por escaneo."
            )

            return (
                finales[0]["fila"],
                None,
            )

        siguiente_numero = (
            numero_recuperacion + 1
        )

        paginas = obtener_paginas_numericas(
            pagina
        )

        avanzo = False

        if siguiente_numero in paginas:
            avanzo = ir_a_pagina_numero(
                pagina,
                siguiente_numero,
            )

        if not avanzo:
            siguiente = encontrar_boton_siguiente(
                pagina
            )

            if siguiente is None:
                break

            try:
                siguiente.click(
                    timeout=5000
                )
                _esperar_cambio_firma(
                    pagina,
                    actividad,
                    firma_actual,
                    timeout_ms=1200,
                )
                avanzo = True
            except Exception:
                return (
                    None,
                    "ERROR_RECUPERANDO_PAGINA",
                )

        if not avanzo:
            break

        firma_nueva = obtener_firma_pagina(
            pagina,
            actividad,
        )

        if firma_nueva == firma_actual:
            break

        numero_recuperacion += 1

    return (
        None,
        "ERROR_RECUPERANDO_FILA",
    )'''

NUEVO_RECUPERACION = '''    # La coincidencia YA quedo confirmada arriba. Aqui solo hay que volver
    # a dejar esa fila en pantalla, asi que vale la pena insistir: se
    # reintenta varias veces, con esperas cada vez mayores y probando
    # todos los terminos que produjeron resultados.
    terminos_recuperacion = [termino_encontrado]
    for termino in terminos_busqueda:
        if termino not in terminos_recuperacion:
            terminos_recuperacion.append(termino)

    espera_base = (
        4200
        if actividad["categoria_prizma"] == "CHALLENGE"
        else 2600
    )

    for intento in range(1, 4):

        espera_intento = espera_base + (intento - 1) * 2500

        for termino in terminos_recuperacion:

            resultado_recuperacion = _recuperar_fila_con_filtro(
                pagina,
                actividad,
                buscador,
                termino,
                espera_intento,
            )

            if resultado_recuperacion == "DUPLICADA":
                return (
                    None,
                    "ERROR_ACTIVIDAD_DUPLICADA",
                )

            if resultado_recuperacion is not None:
                return (
                    resultado_recuperacion,
                    None,
                )

        print(
            "Reintentando recuperacion de fila (intento",
            intento,
            "de 3)...",
        )

        # Una pausa antes del siguiente intento le da margen a PRIZMA
        # para terminar de responder.
        try:
            pagina.wait_for_timeout(700 * intento)
        except Exception:
            pass

    return (
        None,
        "ERROR_RECUPERANDO_FILA",
    )'''


# ------------------------------------------------------------------
# 3) Funcion auxiliar nueva
# ------------------------------------------------------------------
ANCLA_AUXILIAR = '''# ============================================================
# BUSCAR ACTIVIDAD
# ============================================================
'''

AUXILIAR = '''def _recuperar_fila_con_filtro(
    pagina,
    actividad,
    buscador,
    termino,
    timeout_ms,
):
    """Vuelve a filtrar y recorre las paginas hasta ver la fila exacta.

    Devuelve el locator de la fila, la cadena "DUPLICADA" si aparece mas
    de una vez, o None si en este intento no se pudo encontrar.

    A diferencia de la version anterior, un fallo al cambiar de pagina no
    aborta todo el proceso: simplemente termina este intento y el que
    llama vuelve a probar.
    """
    try:
        buscador.fill("")
        pagina.wait_for_timeout(400)
        asegurar_pagina_1(pagina)

        # Pantalla SIN filtro. Sirve para saber cuando el filtro nuevo
        # ya se aplico de verdad y no estamos leyendo lo anterior.
        firma_sin_filtro = obtener_firma_pagina(
            pagina,
            actividad,
        )

        buscador.fill(termino)

        _esperar_resultados_busqueda(
            pagina,
            actividad,
            timeout_ms=timeout_ms,
            firma_previa=firma_sin_filtro,
        )

        asegurar_pagina_1(pagina)
    except Exception:
        return None

    firmas_vistas = set()
    numero_pagina = 1

    while numero_pagina <= 50:

        try:
            firma_actual = obtener_firma_pagina(
                pagina,
                actividad,
            )
        except Exception:
            return None

        if firma_actual in firmas_vistas:
            break

        firmas_vistas.add(firma_actual)

        try:
            resultados = analizar_resultados_pagina(
                pagina,
                actividad,
            )
        except Exception:
            return None

        finales = [
            resultado
            for resultado in resultados
            if resultado["coincide"]
        ]

        if len(finales) > 1:
            return "DUPLICADA"

        if len(finales) == 1:
            print("✅ Fila recuperada por escaneo.")
            return finales[0]["fila"]

        siguiente_numero = numero_pagina + 1

        try:
            paginas = obtener_paginas_numericas(pagina)
        except Exception:
            break

        avanzo = False

        if siguiente_numero in paginas:
            try:
                avanzo = ir_a_pagina_numero(
                    pagina,
                    siguiente_numero,
                )
            except Exception:
                avanzo = False

        if not avanzo:
            try:
                siguiente = encontrar_boton_siguiente(pagina)
            except Exception:
                siguiente = None

            if siguiente is None:
                break

            try:
                siguiente.click(timeout=5000)
                _esperar_cambio_firma(
                    pagina,
                    actividad,
                    firma_actual,
                    timeout_ms=1500,
                )
                avanzo = True
            except Exception:
                # No se pudo avanzar. No es un error definitivo:
                # este intento termina y el que llama reintenta.
                break

        if not avanzo:
            break

        try:
            firma_nueva = obtener_firma_pagina(
                pagina,
                actividad,
            )
        except Exception:
            break

        if firma_nueva == firma_actual:
            break

        numero_pagina += 1

    return None


# ============================================================
# BUSCAR ACTIVIDAD
# ============================================================
'''


CAMBIOS = [
    ("espera de resultados sin datos viejos", VIEJO_ESPERA, NUEVO_ESPERA),
    ("funcion auxiliar de recuperacion", ANCLA_AUXILIAR, AUXILIAR),
    ("recuperacion de fila con reintentos", VIEJO_RECUPERACION, NUEVO_RECUPERACION),
]


def main():
    if not os.path.isfile(RUTA):
        print("ERROR: no encuentro motor_prizma.py en esta carpeta.")
        return 1

    with io.open(RUTA, "r", encoding="utf-8", newline="") as archivo:
        crudo = archivo.read()

    usa_crlf = "\r\n" in crudo
    texto = crudo.replace("\r\n", "\n")
    original = texto

    if "_recuperar_fila_con_filtro" in texto:
        print("Este parche ya estaba aplicado. No se toco nada.")
        return 0

    faltantes = []
    for etiqueta, viejo, _ in CAMBIOS:
        veces = texto.count(viejo)
        if veces != 1:
            faltantes.append((etiqueta, veces))

    if faltantes:
        print("ERROR: motor_prizma.py no es el esperado. No se modifico nada.\n")
        for etiqueta, veces in faltantes:
            print("  - " + etiqueta + ": encontrado " + str(veces) + " vez/veces (se esperaba 1)")
        print("\nManda tu motor_prizma.py para revisarlo.")
        return 1

    for etiqueta, viejo, nuevo in CAMBIOS:
        texto = texto.replace(viejo, nuevo, 1)
        print("  aplicado: " + etiqueta)

    respaldo = "motor_prizma.py.bak-antes-recuperar-fila"
    shutil.copy(RUTA, respaldo)
    print("\n  respaldo guardado en: " + respaldo)

    salida = texto.replace("\n", "\r\n") if usa_crlf else texto

    with io.open(RUTA, "w", encoding="utf-8", newline="") as archivo:
        archivo.write(salida)

    import py_compile
    try:
        py_compile.compile(RUTA, doraise=True)
    except py_compile.PyCompileError as problema:
        restaurar = original.replace("\n", "\r\n") if usa_crlf else original
        with io.open(RUTA, "w", encoding="utf-8", newline="") as archivo:
            archivo.write(restaurar)
        print("\nERROR de sintaxis. Se restauro el archivo original.")
        print(problema)
        return 1

    print("  sintaxis verificada: OK")
    print("\nLISTO. Detén uvicorn con Ctrl + C y arráncalo de nuevo.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
