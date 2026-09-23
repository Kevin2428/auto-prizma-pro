import threading
import time

from test_module import web
from test_module.models import DocumentTest, TestDocument


def test_asignatura_elegida_en_primer_test_se_reutiliza_en_los_siguientes(monkeypatch):
    job_id = "job-asignatura-persistente"
    asignatura_matriz = "Asignatura de la matriz"
    asignatura_prizma = "Asignatura correcta en PRIZMA"
    document = TestDocument(
        tests=[
            DocumentTest(title="Test 1", week="Semana 3", cut="Corte 1"),
            DocumentTest(title="Test 2", week="Semana 5", cut="Corte 2"),
            DocumentTest(title="Test 3", week="Semana 7", cut="Corte 3"),
        ],
        source_name="Prueba persistencia",
    )
    cascada = {
        "programa": "Programa X",
        "asignatura": asignatura_matriz,
    }
    web._JOBS[job_id] = {
        "owner": "tester",
        "document": document,
        "cascada_datos": cascada,
    }

    observado = {}

    class FakeAutomator:
        def __init__(self, *args, **kwargs):
            pass

        def ejecutar_cargue(self, **kwargs):
            cascada_recibida = kwargs["cascada"]
            callback = kwargs["solicitar_asignatura_callback"]

            def responder_seleccion():
                # El callback limpia el Event al comenzar; esperamos a que entre en espera.
                time.sleep(0.05)
                sync = web._AUTOMATION_SYNC[job_id]
                sync["seleccion"]["asignatura"] = asignatura_prizma
                sync["event"].set()

            hilo = threading.Thread(target=responder_seleccion, daemon=True)
            hilo.start()
            elegida = callback([asignatura_prizma], asignatura_matriz)
            hilo.join(timeout=1)

            observado["elegida"] = elegida
            observado["asignatura_despues_del_primer_test"] = cascada_recibida.get("asignatura")
            return {"ok": True, "procesados": 3, "total": 3}

    monkeypatch.setattr(web, "PrizmaTestAutomator", FakeAutomator)

    try:
        resultado = web.ejecutar_test_job(
            job_id=job_id,
            usuario_prizma="usuario",
            clave_prizma="clave",
            user="tester",
            simulacion=True,
            cascada_override=cascada,
        )
    finally:
        web._JOBS.pop(job_id, None)
        web._AUTOMATION_JOBS.pop(job_id, None)
        web._AUTOMATION_SYNC.pop(job_id, None)

    assert resultado["ok"] is True
    assert observado["elegida"] == asignatura_prizma
    assert observado["asignatura_despues_del_primer_test"] == asignatura_prizma
