from test_module.automation import normalizar_texto, resolver_variantes_programa


def test_negocios_internacionales_incluye_aliases_prizma():
    variantes = {normalizar_texto(v) for v in resolver_variantes_programa("Negocios Internacionales")}

    assert normalizar_texto("Negocios Internacionales") in variantes
    assert normalizar_texto("Técnica Virtual Profesional en Procesos Logísticos y de Comercio Exterior") in variantes
    assert normalizar_texto("Técnica Vistural Profesional en Procesos Logísticos y de Comercio Exterior") in variantes


def test_equivalencia_negocios_internacionales_es_bidireccional():
    variantes = {
        normalizar_texto(v)
        for v in resolver_variantes_programa(
            "Técnica Virtual Profesional en Procesos Logísticos y de Comercio Exterior"
        )
    }

    assert normalizar_texto("Negocios Internacionales") in variantes


def test_programa_desconocido_se_conserva_sin_relajarlo():
    assert resolver_variantes_programa("Ingeniería de Sistemas") == ["Ingeniería de Sistemas"]


def test_motor_conserva_equivalencias_historicas_de_programa():
    from motor_prizma import obtener_variantes_programa

    casos = {
        "Administración de Empresas": "Tecnología en Gestión Empresarial",
        "Administración Turística y Hotelera": "Tecnología en Gestión Turística y Hoteles",
        "Contaduría Pública": "Técnica Profesional en Procesos Contables",
        "Negocios Internacionales": "Técnica Virtual Profesional en Procesos Logísticos y de Comercio Exterior",
    }

    for origen, alias in casos.items():
        variantes = obtener_variantes_programa(origen)
        assert normalizar_texto(alias) in variantes


def test_motor_acepta_prefijos_uxt_con_y_sin_separadores():
    from motor_prizma import limpiar_prefijo_tecnico_recurso

    esperado = "relacion entre innovacion"
    assert limpiar_prefijo_tecnico_recurso("U1T1 Relación entre innovación.h5p") == esperado
    assert limpiar_prefijo_tecnico_recurso("U1 T1 Relación entre innovación.h5p") == esperado
    assert limpiar_prefijo_tecnico_recurso("U1-T1-Relación entre innovación.h5p") == esperado
    assert limpiar_prefijo_tecnico_recurso("U1 - T1 - Relación entre innovación.h5p") == esperado
