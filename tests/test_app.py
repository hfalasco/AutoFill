"""Smoke test da tela: app.py roda no harness do Streamlit (`AppTest`), sem navegador."""

from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parents[1] / "app.py"


def test_tela_inicial_roda_sem_erro():
    app = AppTest.from_file(str(APP), default_timeout=60).run()
    assert not app.exception
    assert not app.error


def test_barra_lateral_pede_analises_e_pontos():
    app = AppTest.from_file(str(APP), default_timeout=60).run()
    rotulos = [u.label for u in app.sidebar.get("file_uploader")]
    assert "FO.GSM.LMI.03.08" in rotulos[0]
    assert "Vivência" in rotulos[1]
    assert any("631, 634, 665" in (t.placeholder or "") for t in app.sidebar.text_area)
