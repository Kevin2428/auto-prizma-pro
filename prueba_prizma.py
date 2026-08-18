from playwright.sync_api import sync_playwright


URL_PRIZMA = "https://admin.prizma.site/inicio-sesion"

VERSION_SCRIPT = "PRUEBA_H5P_NUEVO_GESTION_FINANCIERA_V2"

BUILD_INTERNO = "INTERFAZ_WEB_CONEXION_PRIZMA_01"


def probar_conexion_prizma():

    print()
    print("======================================")
    print("PRUEBA DE CONEXIÓN DESDE INTERFAZ")
    print("======================================")

    print(
        "VERSION_SCRIPT =",
        VERSION_SCRIPT
    )

    print(
        "BUILD_INTERNO =",
        BUILD_INTERNO
    )

    print(
        "URL =",
        URL_PRIZMA
    )

    with sync_playwright() as p:

        navegador = p.chromium.launch(
            headless=False
        )

        pagina = navegador.new_page(
            viewport={
                "width": 1440,
                "height": 900
            }
        )

        print()
        print(
            "Abriendo PRIZMA real..."
        )

        pagina.goto(
            URL_PRIZMA
        )

        print()
        print(
            "✅ Navegador abierto."
        )

        print(
            "Inicia sesión manualmente en PRIZMA."
        )

        try:

            pagina.get_by_role(
                "button",
                name="Actividades"
            ).wait_for(
                state="visible",
                timeout=120000
            )

            print()
            print(
                "✅ LOGIN DETECTADO."
            )

            print(
                "La interfaz ya puede controlar "
                "PRIZMA correctamente."
            )

            print()
            print(
                "Esta prueba NO modificará ninguna actividad."
            )

            pagina.wait_for_timeout(
                5000
            )

        except Exception as e:

            print()
            print(
                "❌ No se pudo confirmar el login."
            )

            print(
                e
            )

        print()
        print(
            "Cerrando navegador de prueba..."
        )

        navegador.close()


if __name__ == "__main__":

    probar_conexion_prizma()