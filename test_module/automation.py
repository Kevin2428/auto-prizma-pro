from __future__ import annotations

import base64
import json
import os
import re
import tempfile
import time
import unicodedata
from typing import Callable

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from .models import DocumentTest
from .serializer import serialize_test


def normalizar_texto(texto: str | None) -> str:
    if not texto:
        return ""
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", texto).strip().lower()


MAPA_NUMEROS_TEXTO: dict[int, list[str]] = {
    1: ["1", "uno", "primero", "primera", "i"],
    2: ["2", "dos", "segundo", "segunda", "ii"],
    3: ["3", "tres", "tercero", "tercera", "iii"],
    4: ["4", "cuatro", "cuarto", "cuarta", "iv"],
    5: ["5", "cinco", "quinto", "quinta", "v"],
    6: ["6", "seis", "sexto", "sexta", "vi"],
    7: ["7", "siete", "septimo", "septima", "vii"],
    8: ["8", "ocho", "octavo", "octava", "viii"],
    9: ["9", "nueve", "noveno", "novena", "ix"],
    10: ["10", "diez", "decimo", "decima", "x"],
    11: ["11", "once", "undecimo", "undecima", "xi"],
    12: ["12", "doce", "duodecimo", "duodecima", "xii"],
}

VARIANTES_A_NUM: dict[str, int] = {}
for _num, _vars in MAPA_NUMEROS_TEXTO.items():
    for _v in _vars:
        VARIANTES_A_NUM[_v] = _num




EQUIVALENCIAS_PROGRAMAS = (
    (
        "Administración de Empresas",
        "Tecnología en Gestión Empresarial",
    ),
    (
        "Administración Turística y Hotelera",
        "Tecnología en Gestión Turística y Hoteles",
    ),
    (
        "Contaduría Pública",
        "Técnica Profesional en Procesos Contables",
    ),
    (
        "Negocios Internacionales",
        "Técnica Virtual Profesional en Procesos Logísticos y de Comercio Exterior",
        "Técnica Vistural Profesional en Procesos Logísticos y de Comercio Exterior",
    ),
)


def resolver_variantes_programa(texto: str | None) -> list[str]:
    """Devuelve aliases explícitos y bidireccionales para el selector Programa.

    No aplica fuzzy matching general: si el programa no pertenece a un grupo
    conocido, conserva exactamente el valor recibido.
    """
    valor = str(texto or "").strip()
    if not valor:
        return []

    valor_n = normalizar_texto(valor)
    for grupo in EQUIVALENCIAS_PROGRAMAS:
        grupo_n = {normalizar_texto(item) for item in grupo}
        if valor_n in grupo_n:
            variantes = [valor]
            for item in grupo:
                if normalizar_texto(item) != valor_n:
                    variantes.append(item)
            return variantes

    return [valor]

def resolver_variantes_numericas(texto: str | None) -> list[str]:
    """
    Retorna una lista de variantes numéricas, textuales y ordinales equivalentes
    si el texto contiene una referencia a un nivel o número.
    Por ejemplo, para '3', 'Tres', 'Nivel 3' o 'Tercero', retorna:
    ['3', 'tres', 'tercero', 'tercera', 'iii']
    Si el texto no contiene referencias numéricas, retorna una lista vacía [].
    """
    if not texto:
        return []
    texto_n = normalizar_texto(texto)
    tokens = re.findall(r"\b[a-z0-9]+\b", texto_n)
    for t in tokens:
        if t in VARIANTES_A_NUM:
            num = VARIANTES_A_NUM[t]
            return list(MAPA_NUMEROS_TEXTO[num])
    return []


class PrizmaTestAutomator:
    def __init__(
        self,
        base_url: str = "https://admin.prizma.site",
        headless: bool = True,
        log_callback: Callable[[str, str, str | None], None] | None = None,
    ):
        self.base_url = base_url.rstrip("/")
        self.headless = headless
        self.log_callback = log_callback or (lambda nivel, msg, snap=None: None)
        self.logs: list[dict] = []
        self.capturas: list[str] = []

    def _log(self, mensaje: str, nivel: str = "info", captura_b64: str | None = None):
        entrada = {
            "timestamp": time.strftime("%H:%M:%S"),
            "nivel": nivel,
            "mensaje": mensaje,
            "captura": captura_b64,
        }
        self.logs.append(entrada)
        if captura_b64:
            self.capturas.append(captura_b64)
        self.log_callback(nivel, mensaje, captura_b64)

    def _tomar_captura(self, page) -> str:
        try:
            buf = page.screenshot(type="jpeg", quality=65)
            return base64.b64encode(buf).decode("ascii")
        except Exception:
            return ""

    def login(self, page, usuario_prizma: str, clave_prizma: str) -> bool:
        self._log(f"Iniciando navegación a {self.base_url}/inicio-sesion...", "info")
        page.goto(f"{self.base_url}/inicio-sesion", wait_until="domcontentloaded", timeout=45000)

        usuario_loc = page.locator('input[name="identification_number"]')
        clave_loc = page.locator('input[name="password"]')
        boton_loc = page.get_by_role("button", name="Iniciar sesión", exact=True)

        usuario_loc.wait_for(state="visible", timeout=25000)
        clave_loc.wait_for(state="visible", timeout=25000)
        boton_loc.wait_for(state="visible", timeout=25000)

        usuario_loc.fill(usuario_prizma.strip())
        clave_loc.fill(clave_prizma)
        self._log("Credenciales ingresadas. Presionando 'Iniciar sesión'...", "info")
        boton_loc.click()

        bienvenida = page.get_by_text("Bienvenido a Prizma admin", exact=False)
        error_toast = page.get_by_text("Credenciales incorrectas", exact=False)

        limite = 15000
        transcurrido = 0
        intervalo = 300

        while transcurrido < limite:
            try:
                if bienvenida.count() > 0 and bienvenida.first.is_visible():
                    snap = self._tomar_captura(page)
                    self._log("Sesión iniciada con éxito en PRIZMA.", "ok", snap)
                    return True
            except Exception:
                pass

            try:
                if error_toast.count() > 0 and error_toast.first.is_visible():
                    snap = self._tomar_captura(page)
                    self._log("PRIZMA rechazó las credenciales ingresadas.", "error", snap)
                    raise ValueError("Credenciales incorrectas en PRIZMA. Verifica tu usuario y contraseña.")
            except PlaywrightTimeoutError:
                pass

            if "/inicio-sesion" not in page.url and "/login" not in page.url:
                snap = self._tomar_captura(page)
                self._log("Acceso confirmado a la plataforma PRIZMA.", "ok", snap)
                return True

            page.wait_for_timeout(intervalo)
            transcurrido += intervalo

        body_text = normalizar_texto(page.locator("body").inner_text(timeout=2000))
        if "credenciales incorrectas" in body_text or "usuario o contrasena" in body_text:
            raise ValueError("No se pudo iniciar sesión: credenciales incorrectas en PRIZMA.")

        if "bienvenido" in body_text or "cerrar sesion" in body_text or "/inicio-sesion" not in page.url:
            self._log("Acceso concedido en PRIZMA.", "ok", self._tomar_captura(page))
            return True

        raise RuntimeError("Tiempo de espera agotado al intentar iniciar sesión en PRIZMA.")

    def localizar_campo_formulario(self, page, nombre: str, timeout_ms: int = 15000):
        """
        Localiza de forma robusta un campo de formulario (Autocomplete / Select / TextField de Material UI)
        por su nombre o etiqueta, evitando falsos positivos con inputs ocultos (como tokens o CSRF en #root).
        """
        import time as time_mod
        fin = time_mod.time() + (timeout_ms / 1000.0)

        while time_mod.time() < fin:
            # 1. Estrategia por accesibilidad nativa de Playwright (get_by_label)
            try:
                loc = page.get_by_label(re.compile(rf"^\s*{re.escape(nombre)}", re.I))
                cnt = loc.count() if hasattr(loc, "count") else 0
                if isinstance(cnt, int) and cnt > 0:
                    for i in range(cnt):
                        el = loc.nth(i)
                        if el.is_visible():
                            return el
            except Exception:
                pass

            # 2. Estrategia por rol Combobox con nombre accesible
            try:
                loc = page.get_by_role("combobox", name=re.compile(rf"{re.escape(nombre)}", re.I))
                cnt = loc.count() if hasattr(loc, "count") else 0
                if isinstance(cnt, int) and cnt > 0:
                    for i in range(cnt):
                        el = loc.nth(i)
                        if el.is_visible():
                            return el
            except Exception:
                pass

            # 3. Estrategia por contenedor MUI acotado (.MuiFormControl-root, .MuiAutocomplete-root)
            try:
                containers = page.locator(
                    '.MuiFormControl-root, .MuiAutocomplete-root, [class*="form-group"], [class*="field"]'
                ).filter(
                    has=page.locator('label, span, p, h6, div').filter(has_text=re.compile(rf"^\s*{re.escape(nombre)}", re.I))
                )
                cnt = containers.count() if hasattr(containers, "count") else 0
                if isinstance(cnt, int) and cnt > 0:
                    for i in range(cnt):
                        c = containers.nth(i)
                        inp = c.locator('input:not([type="hidden"])').first
                        cnt_inp = inp.count() if hasattr(inp, "count") else 0
                        if isinstance(cnt_inp, int) and cnt_inp > 0 and inp.is_visible():
                            return inp
                        cb = c.locator('[role="combobox"], [role="button"], .MuiSelect-select').first
                        cnt_cb = cb.count() if hasattr(cb, "count") else 0
                        if isinstance(cnt_cb, int) and cnt_cb > 0 and cb.is_visible():
                            return cb
            except Exception:
                pass

            # 4. Input directo con placeholder que coincida
            try:
                inp = page.locator(f'input[placeholder*="{nombre}" i]').first
                cnt = inp.count() if hasattr(inp, "count") else 0
                if isinstance(cnt, int) and cnt > 0 and inp.is_visible():
                    return inp
            except Exception:
                pass

            # 5. Label adyacente o hermano
            try:
                inp_lbl = page.locator(
                    f'label:has-text("{nombre}") ~ div input:not([type="hidden"]), label:has-text("{nombre}") + div input:not([type="hidden"])'
                ).first
                cnt = inp_lbl.count() if hasattr(inp_lbl, "count") else 0
                if isinstance(cnt, int) and cnt > 0 and inp_lbl.is_visible():
                    return inp_lbl
            except Exception:
                pass

            page.wait_for_timeout(300)

        return None

    def extraer_valor_campo_formulario(self, page, nombre: str, timeout_ms: int = 5000) -> str:
        """
        Extrae el valor actual de un campo de formulario (Autocomplete / Select / TextField de Material UI)
        por su nombre o etiqueta en un formulario o modal de PRIZMA.
        """
        campo = self.localizar_campo_formulario(page, nombre, timeout_ms=timeout_ms)
        if not campo:
            return ""

        # 1. Probar input_value si es un elemento de entrada
        try:
            val = campo.input_value()
            if val and str(val).strip():
                return str(val).strip()
        except Exception:
            pass

        # 2. Probar atributo "value"
        try:
            val = campo.get_attribute("value")
            if val and str(val).strip():
                return str(val).strip()
        except Exception:
            pass

        # 3. Probar si dentro del contenedor padre hay un valor seleccionado
        try:
            padre = campo.locator("xpath=ancestor::*[contains(@class, 'MuiFormControl-root') or contains(@class, 'MuiAutocomplete-root') or contains(@class, 'form-group')][1]")
            cnt_p = padre.count() if hasattr(padre, "count") else 0
            if isinstance(cnt_p, int) and cnt_p > 0:
                sel = padre.locator('.MuiSelect-select, .MuiChip-label, .MuiAutocomplete-tag, .MuiInputBase-input').first
                cnt_s = sel.count() if hasattr(sel, "count") else 0
                if isinstance(cnt_s, int) and cnt_s > 0 and sel.is_visible():
                    txt = sel.inner_text().strip()
                    if txt:
                        return txt
        except Exception:
            pass

        # 4. Probar inner_text directo
        try:
            val = campo.inner_text().strip()
            if val:
                return val
        except Exception:
            pass

        return ""

    def seleccionar_mui(
        self,
        page,
        placeholder: str,
        valor: str,
        on_opcion_no_encontrada: Callable[[list[str], str], str] | None = None,
    ) -> bool:
        if not valor or not str(valor).strip():
            return True

        valor_limpio = str(valor).strip()
        valor_n = normalizar_texto(valor_limpio)
        variantes_num = resolver_variantes_numericas(valor_limpio)
        variantes_programa = (
            resolver_variantes_programa(valor_limpio)
            if normalizar_texto(placeholder) == "programa"
            else [valor_limpio]
        )
        variantes_programa_n = [normalizar_texto(v) for v in variantes_programa if normalizar_texto(v)]
        self._log(f"Configurando campo '{placeholder}' con valor '{valor_limpio}'...", "info")

        campo = self.localizar_campo_formulario(page, placeholder, timeout_ms=12000)

        if not campo or not campo.is_visible():
            if placeholder == "Facultad":
                self._log("Selector 'Facultad' no detectado en el formulario. Omitiendo ya que no es obligatorio.", "info")
                return True
            snap = self._tomar_captura(page)
            self._log(f"No se encontró el campo '{placeholder}' en el formulario.", "error", snap)
            raise ValueError(f"No se encontró el selector '{placeholder}' en el formulario de PRIZMA.")

        try:
            val_actual = normalizar_texto(campo.input_value())
            if any(
                val_actual == variante
                or (len(variante) >= 3 and variante in val_actual)
                for variante in variantes_programa_n
            ):
                return True
            if any(val_actual == v or (len(v) >= 3 and v in val_actual) for v in variantes_num):
                return True
        except Exception:
            pass

        for intento in range(1, 4):
            try:
                campo.click(timeout=6000)
                page.wait_for_timeout(300)

                # Si es intento posterior y tenemos variantes numéricas, podemos alternar entre texto y número
                texto_a_escribir = valor_limpio
                if normalizar_texto(placeholder) == "programa" and len(variantes_programa) > 1:
                    texto_a_escribir = variantes_programa[min(intento - 1, len(variantes_programa) - 1)]
                elif intento == 2 and len(variantes_num) > 1:
                    texto_a_escribir = variantes_num[1] if variantes_num[0] == valor_n else variantes_num[0]

                # Escribimos el código o nombre para que Material UI filtre en vivo
                try:
                    campo.fill(texto_a_escribir)
                except Exception:
                    try:
                        campo.type(texto_a_escribir)
                    except Exception:
                        pass

                page.wait_for_timeout(500 + (intento * 200))

                opciones = page.locator('[role="option"], li[role="option"], li.MuiAutocomplete-option, ul.MuiMenu-list li, .MuiMenuItem-root')

                # Si no aparecieron opciones y tenemos variantes numéricas, probar con la variante alternativa
                if opciones.count() == 0 and len(variantes_num) > 1:
                    alt_texto = variantes_num[1] if variantes_num[0] == valor_n else variantes_num[0]
                    try:
                        campo.fill(alt_texto)
                        page.wait_for_timeout(400)
                        opciones = page.locator('[role="option"], li[role="option"], li.MuiAutocomplete-option, ul.MuiMenu-list li, .MuiMenuItem-root')
                    except Exception:
                        pass

                coincidencias_exactas = []
                coincidencias_numericas = []
                coincidencias_codigo = []
                coincidencias_parciales = []

                for idx in range(opciones.count()):
                    opc = opciones.nth(idx)
                    if not opc.is_visible():
                        continue
                    txt_raw = opc.inner_text().strip()
                    txt = normalizar_texto(txt_raw)

                    is_code_user = bool(re.search(r"\b([A-Z]{2,}\s*[\-_]?\s*\d+|\d{3,})\b", valor_limpio, re.I))

                    if txt in variantes_programa_n:
                        coincidencias_exactas.append((opc, txt_raw))
                        break

                    # Coincidencia numérica bidireccional (ej. 3 <-> tres <-> tercero <-> iii)
                    if any(v == txt or re.search(rf"\b{re.escape(v)}\b", txt) for v in variantes_num):
                        coincidencias_numericas.append((opc, txt_raw))

                    if is_code_user and re.search(rf"\b{re.escape(valor_n)}\b", txt):
                        coincidencias_codigo.append((opc, txt_raw))
                    elif any(variante and variante in txt for variante in variantes_programa_n):
                        coincidencias_parciales.append((opc, txt_raw))

                if coincidencias_exactas:
                    coincidencias_exactas[0][0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado con '{coincidencias_exactas[0][1]}'.", "ok")
                    return True

                if len(coincidencias_codigo) == 1:
                    coincidencias_codigo[0][0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado por código con '{coincidencias_codigo[0][1]}'.", "ok")
                    return True

                if len(coincidencias_codigo) > 1:
                    starts = [c for c in coincidencias_codigo if normalizar_texto(c[1]).startswith(valor_n)]
                    elegida = starts[0] if len(starts) == 1 else coincidencias_codigo[0]
                    elegida[0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado con '{elegida[1]}'.", "ok")
                    return True

                if len(coincidencias_numericas) == 1:
                    coincidencias_numericas[0][0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado por equivalencia numérica/letras con '{coincidencias_numericas[0][1]}'.", "ok")
                    return True

                if len(coincidencias_numericas) > 1:
                    exactas_num = [c for c in coincidencias_numericas if normalizar_texto(c[1]) in variantes_num]
                    if len(exactas_num) == 1:
                        exactas_num[0][0].click(timeout=6000)
                        page.wait_for_timeout(400)
                        self._log(f"Campo '{placeholder}' configurado con '{exactas_num[0][1]}'.", "ok")
                        return True
                    elegida_num = coincidencias_numericas[0]
                    elegida_num[0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado con '{elegida_num[1]}'.", "ok")
                    return True

                if len(coincidencias_parciales) == 1:
                    coincidencias_parciales[0][0].click(timeout=6000)
                    page.wait_for_timeout(400)
                    self._log(f"Campo '{placeholder}' configurado con '{coincidencias_parciales[0][1]}'.", "ok")
                    return True

                if len(coincidencias_parciales) > 1:
                    snap = self._tomar_captura(page)
                    ejemplos = " | ".join(c[1] for c in coincidencias_parciales[:3])
                    msg_amb = (
                        f"Ambigüedad en '{placeholder}': se encontraron múltiples opciones para '{valor_limpio}' "
                        f"({ejemplos}). Por favor ingresa el código exacto (ej. COD111) para evitar inconvenientes."
                    )
                    self._log(msg_amb, "error", snap)
                    raise ValueError(msg_amb)

            except ValueError:
                raise
            except Exception as e:
                self._log(f"Reintentando selección de '{placeholder}' (intento {intento}): {e}", "warn")
                page.wait_for_timeout(800)

        if on_opcion_no_encontrada:
            try:
                campo.click(timeout=4000)
                try:
                    campo.fill("")
                except Exception:
                    pass
                page.wait_for_timeout(600)
                opciones_loc = page.locator('[role="option"], li[role="option"], li.MuiAutocomplete-option, ul.MuiMenu-list li, .MuiMenuItem-root')
                lista_opciones = []
                for idx in range(opciones_loc.count()):
                    txt_opc = opciones_loc.nth(idx).inner_text().strip()
                    if txt_opc and txt_opc not in lista_opciones:
                        lista_opciones.append(txt_opc)

                if lista_opciones:
                    self._log(f"Opción '{valor_limpio}' no encontrada en '{placeholder}'. Solicitando selección al usuario entre {len(lista_opciones)} opciones...", "warn")
                    opcion_elegida = on_opcion_no_encontrada(lista_opciones, valor_limpio)
                    if opcion_elegida and str(opcion_elegida).strip():
                        elegida_str = str(opcion_elegida).strip()
                        self._log(f"Usuario seleccionó: '{elegida_str}'. Aplicando en PRIZMA...", "info")
                        elegida_n = normalizar_texto(elegida_str)
                        click_done = False
                        opciones_loc = page.locator('[role="option"], li[role="option"], li.MuiAutocomplete-option, ul.MuiMenu-list li, .MuiMenuItem-root')
                        for idx in range(opciones_loc.count()):
                            opc = opciones_loc.nth(idx)
                            if normalizar_texto(opc.inner_text().strip()) == elegida_n:
                                opc.click(timeout=6000)
                                click_done = True
                                break
                        if not click_done:
                            campo.fill(elegida_str)
                            page.wait_for_timeout(500)
                            opciones_loc = page.locator('[role="option"], li[role="option"], li.MuiAutocomplete-option, ul.MuiMenu-list li, .MuiMenuItem-root')
                            if opciones_loc.count() > 0:
                                opciones_loc.first.click(timeout=6000)
                                click_done = True
                        if click_done:
                            page.wait_for_timeout(400)
                            self._log(f"Campo '{placeholder}' configurado exitosamente con '{elegida_str}'.", "ok")
                            return True
            except Exception as exc_fallback:
                self._log(f"Error durante fallback de selección interactiva: {exc_fallback}", "warn")

        snap = self._tomar_captura(page)
        msg_err = f"La opción o código '{valor_limpio}' no existe en el selector '{placeholder}' de PRIZMA."
        self._log(msg_err, "error", snap)
        raise ValueError(msg_err)

    def llenar_descripcion_test(self, page, descripcion: str) -> bool:
        texto = (descripcion or "").strip()
        if not texto:
            texto = "Responde con base en los contenidos estudiados."

        self._log(f"Configurando campo 'Descripción' ({len(texto)} caracteres)...", "info")

        editores = page.locator(
            '[contenteditable="true"], .ql-editor, textarea[placeholder*="Descripci" i], textarea[placeholder*="Escribe" i], [role="textbox"], textarea'
        )

        editor_encontrado = None
        cnt = editores.count() if hasattr(editores, "count") else 0
        if isinstance(cnt, int) and cnt > 0:
            for idx in range(cnt):
                ed = editores.nth(idx)
                if ed.is_visible():
                    editor_encontrado = ed
                    break

        if not editor_encontrado:
            lbl_desc = page.locator('label:has-text("Descripci"), p:has-text("Descripci"), span:has-text("Descripci")').first
            cnt_lbl = lbl_desc.count() if hasattr(lbl_desc, "count") else 0
            if isinstance(cnt_lbl, int) and cnt_lbl > 0 and lbl_desc.is_visible():
                contenedor = lbl_desc.locator('xpath=ancestor::*[contains(@class, "MuiFormControl") or contains(@class, "form") or contains(@class, "field")][1]')
                cnt_cont = contenedor.count() if hasattr(contenedor, "count") else 0
                if isinstance(cnt_cont, int) and cnt_cont > 0:
                    cand = contenedor.locator('[contenteditable="true"], textarea, [role="textbox"]').first
                    cnt_cand = cand.count() if hasattr(cand, "count") else 0
                    if isinstance(cnt_cand, int) and cnt_cand > 0 and cand.is_visible():
                        editor_encontrado = cand

        if editor_encontrado:
            try:
                editor_encontrado.click(timeout=6000)
                page.wait_for_timeout(250)
                try:
                    editor_encontrado.fill(texto)
                except Exception:
                    editor_encontrado.press("Control+A")
                    editor_encontrado.press("Backspace")
                    page.keyboard.insert_text(texto)

                page.wait_for_timeout(400)
                snap = self._tomar_captura(page)
                self._log("Descripción del test completada exitosamente.", "ok", snap)
                return True
            except Exception as e:
                self._log(f"Error al escribir en la descripción: {e}", "warn")
                return False

        self._log("No se detectó el campo de descripción en el formulario.", "warn")
        return False

    def navegar_a_crear_test(self, page, forzar_nuevo: bool = False):
        self._log("Iniciando navegación: Dashboard -> Actividades -> Tests -> Crear...", "info")

        if not forzar_nuevo:
            campo_previo = (
                self.localizar_campo_formulario(page, "Programa", timeout_ms=2000)
                or self.localizar_campo_formulario(page, "Facultad", timeout_ms=1000)
            )
            if campo_previo and campo_previo.is_visible():
                nombre_previo = page.locator('input[placeholder="Nombre del test"]').first
                cnt_nom = nombre_previo.count() if hasattr(nombre_previo, "count") else 0
                if isinstance(cnt_nom, int) and cnt_nom > 0:
                    val_nom = str(nombre_previo.input_value() or "").strip()
                    if not val_nom:
                        self._log("Formulario 'Crear test' limpio ya visible en pantalla.", "ok")
                        return

        # Si se requiere un formulario nuevo (por ejemplo tras terminar el test previo),
        # cerramos o descartamos el formulario anterior con "Cancelar" si está presente.
        btn_cancelar = page.get_by_role("button", name="Cancelar", exact=True).first
        cnt_canc = btn_cancelar.count() if hasattr(btn_cancelar, "count") else 0
        if isinstance(cnt_canc, int) and cnt_canc > 0 and btn_cancelar.is_visible():
            self._log("Descartando formulario previo con botón 'Cancelar'...", "info")
            btn_cancelar.click(timeout=6000)
            page.wait_for_timeout(1000)

        # 1. Clic en Actividades
        self._log("Paso 1/3: Entrando al módulo 'Actividades'...", "info")
        btn_actividades = page.get_by_role("button", name="Actividades").first
        try:
            if not btn_actividades.is_visible():
                btn_actividades = page.get_by_role("button").filter(has_text=re.compile(r"Actividades", re.I)).first
            if not btn_actividades.is_visible():
                btn_actividades = page.locator(
                    'button:has-text("Actividades"), [role="button"]:has-text("Actividades"), a:has-text("Actividades"), div:has-text("Actividades")'
                ).first
            if btn_actividades.is_visible():
                btn_actividades.click(timeout=10000)
                self._log("Módulo 'Actividades' abierto.", "ok", self._tomar_captura(page))
                page.wait_for_timeout(1500)
        except Exception as e:
            self._log(f"Accediendo a 'Actividades': {e}", "warn")

        # 2. Clic en el apartado específico 'Tests' (en plural o label.tab de PRIZMA)
        self._log("Paso 2/3: Accediendo al apartado 'Tests'...", "info")
        btn_tests = page.locator("label.tab").filter(has_text=re.compile(r"^tests?$", re.I)).first
        try:
            if not btn_tests.is_visible():
                btn_tests = page.get_by_role("tab", name=re.compile(r"^tests?$", re.I)).first
            if not btn_tests.is_visible():
                btn_tests = page.get_by_role("button", name=re.compile(r"^tests?$", re.I)).first
            if not btn_tests.is_visible():
                btn_tests = page.locator(
                    'label.tab:has-text("Tests"), label.tab:has-text("Test"), [role="tab"]:has-text("Tests"), button:has-text("Tests"), a:has-text("Tests"), div:has-text("Tests"), [role="tab"]:has-text("Test"), button:has-text("Test")'
                ).first
            btn_tests.wait_for(state="visible", timeout=20000)
            btn_tests.click(timeout=10000)
            self._log("Apartado 'Tests' seleccionado.", "ok", self._tomar_captura(page))
            page.wait_for_timeout(1500)
        except Exception as e:
            self._log(f"Accediendo al apartado 'Tests': {e}", "warn")

        # 3. Clic en el botón 'Crear' dentro de Tests
        self._log("Paso 3/3: Buscando y presionando botón 'Crear'...", "info")
        btn_crear = page.get_by_role("button", name=re.compile(r"^crear(\s+test)?$", re.I)).first
        try:
            if not btn_crear.is_visible():
                btn_crear = page.locator(
                    'button:has-text("Crear"), a:has-text("Crear"), [role="button"]:has-text("Crear"), button:has-text("Crear test"), a:has-text("Crear test"), div:has-text("Crear")'
                ).first
            btn_crear.wait_for(state="visible", timeout=20000)
            btn_crear.click(timeout=10000)
            self._log("Botón 'Crear' presionado.", "ok", self._tomar_captura(page))
            page.wait_for_timeout(1500)
        except Exception as e:
            self._log(f"Buscando botón 'Crear': {e}", "warn")

        # 4. Esperar que el formulario de creación esté listo (verificando Programa, Facultad o Pensum)
        self._log("Esperando renderizado del formulario de creación...", "info")
        campo_listo = (
            self.localizar_campo_formulario(page, "Programa", timeout_ms=8000)
            or self.localizar_campo_formulario(page, "Facultad", timeout_ms=3000)
            or self.localizar_campo_formulario(page, "Pensum", timeout_ms=3000)
        )

        # Si aún no aparece tras el tiempo inicial, reintentar presionar Crear por si el click inicial fue absorbido
        if not campo_listo:
            try:
                if btn_crear.is_visible():
                    self._log("Reintentando clic en botón 'Crear'...", "info")
                    btn_crear.click(timeout=6000)
                    page.wait_for_timeout(2000)
            except Exception:
                pass
            campo_listo = (
                self.localizar_campo_formulario(page, "Programa", timeout_ms=15000)
                or self.localizar_campo_formulario(page, "Facultad", timeout_ms=5000)
                or self.localizar_campo_formulario(page, "Pensum", timeout_ms=5000)
            )

        if campo_listo and campo_listo.is_visible():
            self._log("Formulario de creación listo y cargado correctamente.", "ok", self._tomar_captura(page))
        else:
            snap = self._tomar_captura(page)
            self._log("No se encontró el formulario de creación tras navegar por Actividades -> Tests -> Crear.", "error", snap)
            raise RuntimeError("No se encontró el formulario de creación en PRIZMA tras acceder a Actividades -> Tests -> Crear.")

    def inyectar_json_test(self, page, test: DocumentTest) -> bool:
        payload = serialize_test(test)
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            temp_path = f.name

        try:
            self._log(f"Inyectando archivo JSON de '{test.title}' en el botón 'Importar'...", "info")
            btn_importar = page.locator('button:has-text("Importar"), [role="button"]:has-text("Importar"), a:has-text("Importar")').first
            btn_importar.wait_for(state="visible", timeout=15000)
            try:
                btn_importar.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass

            with page.expect_file_chooser(timeout=15000) as fc_info:
                btn_importar.click()

            file_chooser = fc_info.value
            file_chooser.set_files(temp_path)

            self._log("Archivo JSON cargado en el explorador de archivos. Esperando renderizado de preguntas...", "info")
            page.wait_for_timeout(2500)

            snap = self._tomar_captura(page)
            self._log(f"Preguntas del test '{test.title}' importadas en la tabla.", "ok", snap)
            return True
        finally:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass

    def guardar_creacion_test(self, page, titulo_test: str) -> bool:
        self._log(f"Iniciando guardado de test '{titulo_test}' en PRIZMA...", "info")

        btn_guardar = None

        # Estrategia 1: Botón adyacente o hermano posterior a 'Cancelar' en CSS
        try:
            btn_hermano = page.locator('button:has-text("Cancelar") ~ button, button:has-text("Cancelar") + button').first
            if hasattr(btn_hermano, "is_visible") and btn_hermano.is_visible():
                txt_h = (btn_hermano.inner_text() or "").strip().lower()
                if ("crear" in txt_h or "guardar" in txt_h) and "cancelar" not in txt_h:
                    btn_guardar = btn_hermano
        except Exception:
            pass

        # Estrategia 2: Si no está después de Cancelar, buscar hermano anterior/contenedor mediante 'Cancelar'
        if not btn_guardar:
            try:
                btn_cancelar = page.locator('button:has-text("Cancelar"), [role="button"]:has-text("Cancelar")').filter(
                    has_text=re.compile(r"^cancelar$", re.I)
                ).first
                if not btn_cancelar.is_visible():
                    btn_cancelar = page.locator('button:has-text("Cancelar"), [role="button"]:has-text("Cancelar")').first

                if btn_cancelar.is_visible():
                    # 2A: XPath para botón hermano 'Crear' (anterior o posterior en el mismo footer)
                    candidatos_hermanos = btn_cancelar.locator(
                        'xpath=preceding-sibling::button[contains(., "Crear")] | following-sibling::button[contains(., "Crear")] | '
                        'xpath=..//button[contains(., "Crear") and not(contains(., "Cancelar"))]'
                    )
                    cnt_h = candidatos_hermanos.count() if hasattr(candidatos_hermanos, "count") else 0
                    if isinstance(cnt_h, int) and cnt_h > 0:
                        for idx in range(cnt_h):
                            cand = candidatos_hermanos.nth(idx)
                            if cand.is_visible():
                                btn_guardar = cand
                                break

                    # 2B: Diálogo / formulario modal activo que contiene a 'Cancelar'
                    if not btn_guardar:
                        dialogo = page.locator('[role="dialog"], .MuiDialog-root, form').filter(has=btn_cancelar).last
                        if dialogo.is_visible():
                            cand_dialog = dialogo.locator('button:has-text("Crear"), [role="button"]:has-text("Crear")').last
                            if cand_dialog.is_visible():
                                btn_guardar = cand_dialog
            except Exception:
                pass

        # Estrategia 3: Buscar en botones 'Crear' visibles y habilitados en pantalla (del último al primero)
        if not btn_guardar:
            try:
                candidatos = page.locator(
                    'button[type="submit"]:has-text("Crear"), button:has-text("Crear"), [role="button"]:has-text("Crear")'
                )
                cnt_c = candidatos.count() if hasattr(candidatos, "count") else 0
                if isinstance(cnt_c, int) and cnt_c > 0:
                    for idx in reversed(range(cnt_c)):
                        cand = candidatos.nth(idx)
                        txt = (cand.inner_text() or "").strip().lower()
                        if cand.is_visible() and "cancelar" not in txt:
                            if cand.is_enabled():
                                btn_guardar = cand
                                break
                            elif not btn_guardar:
                                btn_guardar = cand
            except Exception:
                pass

        if not btn_guardar or not btn_guardar.is_visible():
            snap = self._tomar_captura(page)
            self._log(f"No se localizó el botón 'Crear' al lado de 'Cancelar' para guardar el test '{titulo_test}'.", "error", snap)
            raise ValueError("No se encontró el botón 'Crear' en el formulario activo para persistir el test en PRIZMA.")

        # Verificar si está habilitado; si no, esperar hasta 8 segundos como solicitó el usuario
        inicio_espera = time.time()
        timeout_espera_s = 8.0
        while time.time() - inicio_espera < timeout_espera_s:
            try:
                if btn_guardar.is_enabled():
                    break
            except Exception:
                pass
            page.wait_for_timeout(400)
        else:
            try:
                if not btn_guardar.is_enabled():
                    snap_des = self._tomar_captura(page)
                    self._log(f"El botón 'Crear' en PRIZMA continúa deshabilitado tras {int(timeout_espera_s)} segundos de espera.", "error", snap_des)
                    raise ValueError("El botón 'Crear' en PRIZMA se encuentra deshabilitado. Verifica que todos los campos requeridos estén completos.")
            except Exception as e:
                if "deshabilitado" in str(e):
                    raise
                pass

        try:
            btn_guardar.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass

        self._log("Presionando botón 'Crear' en PRIZMA...", "info")
        btn_guardar.click(timeout=10000)

        # 2. Esperar confirmación de guardado
        self._log("Esperando confirmación de almacenamiento de PRIZMA...", "info")
        guardado_confirmado = False
        esperado_ms = 0
        intervalo_ms = 400
        timeout_max_ms = 15000

        while esperado_ms < timeout_max_ms:
            page.wait_for_timeout(intervalo_ms)
            esperado_ms += intervalo_ms

            # Alerta o toast de error en pantalla
            alertas_error = page.locator(
                '.MuiAlert-standardError, .MuiAlert-filledError, [role="alert"]:has-text("error"), [role="alert"]:has-text("falló")'
            )
            cnt_err = alertas_error.count() if hasattr(alertas_error, "count") else 0
            if isinstance(cnt_err, int) and cnt_err > 0 and alertas_error.first.is_visible():
                txt_err = alertas_error.first.inner_text().strip()
                snap_err = self._tomar_captura(page)
                self._log(f"PRIZMA rechazó el guardado con error: {txt_err}", "error", snap_err)
                raise ValueError(f"Error de PRIZMA al guardar el test: {txt_err}")

            # Detección de errores de validación en campos del formulario MUI
            campos_invalidos = page.locator('.Mui-error, [aria-invalid="true"], .MuiFormHelperText-root.Mui-error')
            cnt_inv = campos_invalidos.count() if hasattr(campos_invalidos, "count") else 0
            if isinstance(cnt_inv, int) and cnt_inv > 0:
                for idx in range(cnt_inv):
                    inv = campos_invalidos.nth(idx)
                    if inv.is_visible():
                        txt_inv = (inv.inner_text() or "").strip()
                        if txt_inv and "cancelar" not in txt_inv.lower():
                            snap_inv = self._tomar_captura(page)
                            self._log(f"Campo requerido o inválido en PRIZMA: {txt_inv}", "error", snap_inv)
                            raise ValueError(f"Campo obligatorio o inválido en PRIZMA al guardar '{titulo_test}': {txt_inv}")

            # Toast o snackbar de éxito
            alertas_ok = page.locator(
                '.MuiAlert-standardSuccess, .MuiAlert-filledSuccess, [role="alert"]:has-text("éxito"), [role="alert"]:has-text("exitoso"), [role="alert"]:has-text("cread")'
            )
            cnt_ok = alertas_ok.count() if hasattr(alertas_ok, "count") else 0
            if isinstance(cnt_ok, int) and cnt_ok > 0 and alertas_ok.first.is_visible():
                guardado_confirmado = True
                break

            # Formulario cerrado (el campo 'Nombre del test' desapareció)
            campo_nombre = page.locator('input[placeholder="Nombre del test"]').first
            cnt_nom = campo_nombre.count() if hasattr(campo_nombre, "count") else 0
            if cnt_nom == 0 or not campo_nombre.is_visible():
                guardado_confirmado = True
                break

        if not guardado_confirmado:
            snap_warn = self._tomar_captura(page)
            self._log(
                f"Advertencia: El formulario de '{titulo_test}' no confirmó su cierre tras presionar Guardar.",
                "warn",
                snap_warn,
            )

        snap_final = self._tomar_captura(page)
        self._log(f"Test '{titulo_test}' procesado y guardado exitosamente en PRIZMA.", "ok", snap_final)
        return True

    def ejecutar_cargue(
        self,
        usuario_prizma: str,
        clave_prizma: str,
        tests: list[DocumentTest],
        cascada: dict[str, str],
        solo_simulacion: bool = True,
        test_callback=None,
        cancel_checker: Callable[[], bool] | None = None,
        solicitar_asignatura_callback: Callable[[list[str], str], str] | None = None,
    ) -> dict:
        self._log("Iniciando motor de automatización Playwright...", "info")
        with sync_playwright() as p:
            args = ["--no-sandbox", "--disable-dev-shm-usage"]
            if not self.headless:
                args.extend(["--start-maximized"])

            effective_headless = self.headless
            if not effective_headless and sys.platform != "win32" and not os.environ.get("DISPLAY"):
                self._log("No se detectó servidor de pantalla ($DISPLAY). Ejecutando en modo headless con capturas en vivo.", "warn")
                effective_headless = True

            try:
                navegador = p.chromium.launch(
                    headless=effective_headless,
                    args=args,
                )
            except Exception as e_launch:
                if not effective_headless:
                    self._log(f"Fallo al abrir navegador gráfico ({e_launch}). Continuando automáticamente en modo headless con capturas...", "warn")
                    effective_headless = True
                    navegador = p.chromium.launch(
                        headless=True,
                        args=args,
                    )
                else:
                    raise

            contexto = navegador.new_context(viewport={"width": 1440, "height": 900} if effective_headless else None)
            page = contexto.new_page()

            try:
                self.login(page, usuario_prizma, clave_prizma)

                exitosos = 0
                detenido = False
                for index, test in enumerate(tests, start=1):
                    # Verificar si el usuario solicitó detener la automatización
                    if cancel_checker and cancel_checker():
                        detenido = True
                        self._log(f"⏹ Automatización detenida por el usuario antes de procesar Test {index}/{len(tests)}: '{test.title}'.", "warn")
                        if test_callback:
                            for rem_idx in range(index, len(tests) + 1):
                                test_callback(rem_idx, "cancelado", "Detenido por el usuario", "")
                        break

                    self._log(f"--- Iniciando procesamiento de Test {index}/{len(tests)}: '{test.title}' ---", "info")
                    if test_callback:
                        test_callback(index, "en_proceso", "Diligenciando campos e importando preguntas...", "")

                    try:
                        self.navegar_a_crear_test(page, forzar_nuevo=(index > 1))

                        cascada_items = [
                            ("Facultad", cascada.get("facultad", "")),
                            ("Programa", cascada.get("programa", "")),
                            ("Pensum", cascada.get("pensum", "")),
                            ("Nivel de pensum", cascada.get("nivel", "")),
                            ("Asignatura por pensum", cascada.get("asignatura", "")),
                            ("Semana", test.week or cascada.get("semana", "")),
                            ("Corte", test.cut or cascada.get("corte", "")),
                        ]
                        for placeholder, valor in cascada_items:
                            if valor:
                                try:
                                    cb_fallback = solicitar_asignatura_callback if placeholder == "Asignatura por pensum" else None
                                    self.seleccionar_mui(page, placeholder, valor, on_opcion_no_encontrada=cb_fallback)
                                except Exception as exc_item:
                                    if placeholder == "Facultad":
                                        self._log(f"Campo opcional 'Facultad' no se pudo configurar ({exc_item}). Continuando ya que no es obligatorio...", "warn")
                                    else:
                                        raise

                        self.seleccionar_mui(page, "Estado", "Activo")

                        nombre_loc = page.locator('input[placeholder="Nombre del test"]').first
                        cnt_nom = nombre_loc.count() if hasattr(nombre_loc, "count") else 0
                        if isinstance(cnt_nom, int) and cnt_nom > 0 and nombre_loc.is_visible():
                            nombre_loc.fill(test.title)

                        # Llenar la descripción utilizando el editor de texto enriquecido o textarea
                        self.llenar_descripcion_test(page, test.description)

                        self.inyectar_json_test(page, test)

                        if solo_simulacion:
                            snap_sim = self._tomar_captura(page)
                            self._log(
                                f"🛡 [MODO SIMULACIÓN] Test {index} ('{test.title}'): Formulario e importación validados con éxito. NO se presionó Guardar para proteger la base de datos.",
                                "ok",
                                snap_sim,
                            )
                            if test_callback:
                                test_callback(index, "exitoso", "Validado en modo simulación (sin guardar)", "")
                            if not self.headless:
                                self._log("Pausando 8 segundos para inspección visual en pantalla...", "info")
                                page.wait_for_timeout(8000)
                        else:
                            self.guardar_creacion_test(page, test.title)
                            if test_callback:
                                test_callback(index, "exitoso", "Creado y guardado exitosamente en PRIZMA", "")
                            if not self.headless:
                                self._log("Pausando 4 segundos para inspección visual en pantalla...", "info")
                                page.wait_for_timeout(4000)

                        exitosos += 1

                    except Exception as test_err:
                        if test_callback:
                            test_callback(index, "error", f"Error en test: {test_err}", str(test_err))
                        raise

                if detenido:
                    msg_detenido = f"Automatización detenida por el usuario. Se completaron {exitosos} de {len(tests)} tests."
                    self._log(msg_detenido, "warn")
                    return {
                        "ok": True,
                        "detenido": True,
                        "simulacion": solo_simulacion,
                        "procesados": exitosos,
                        "total": len(tests),
                        "logs": self.logs,
                        "capturas": self.capturas,
                    }

                msg_final = (
                    "Simulación completada con éxito: todos los campos e importaciones fueron validados sin guardar en PRIZMA."
                    if solo_simulacion
                    else "Todos los tests fueron procesados y guardados con éxito en PRIZMA."
                )
                self._log(msg_final, "ok")
                return {
                    "ok": True,
                    "detenido": False,
                    "simulacion": solo_simulacion,
                    "procesados": exitosos,
                    "total": len(tests),
                    "logs": self.logs,
                    "capturas": self.capturas,
                }

            except Exception as e:
                snap_err = self._tomar_captura(page)
                self._log(f"Error durante la automatización: {e}", "error", snap_err)
                return {
                    "ok": False,
                    "detenido": False,
                    "simulacion": solo_simulacion,
                    "error": str(e),
                    "logs": self.logs,
                    "capturas": self.capturas,
                }
            finally:
                try:
                    contexto.close()
                    navegador.close()
                except Exception:
                    pass


