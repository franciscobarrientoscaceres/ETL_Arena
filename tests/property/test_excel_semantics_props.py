"""Property tests de excel_semantics."""

from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.excel_semantics import codigo_falla_excel, redondear_excel, texto_excel


def _find(buscado: str, texto: str, inicio: int = 1):
    """FIND de Excel: posición 1-based o error (None)."""
    p = texto.find(buscado, inicio - 1)
    return None if p == -1 else p + 1


def _mid(texto: str, inicio: int, largo: int):
    """MID de Excel: largo negativo es error (None)."""
    if largo < 0:
        return None
    return texto[inicio - 1 : inicio - 1 + largo]


def _referencia(g) -> str:
    """Traducción literal de IFERROR(MID(G,1,FIND(" ",G,1)-1),CONCATENATE("F",G))."""
    s = texto_excel(g)
    pos = _find(" ", s, 1)
    r = None if pos is None else _mid(s, 1, pos - 1)
    return "F" + s if r is None else r


_descripciones = st.one_of(
    st.none(),
    st.integers(min_value=0, max_value=400).map(float),
    st.text(alphabet=st.sampled_from("F0123456789 ABCXYZ/-"), max_size=30),
    st.sampled_from(["NO FAULTS", "F55 EXTERNAL FAULT/OVGR", "F1 Watchdog", " X", ""]),
)


# Feature: etl-arena-availability, Property 8
@settings(max_examples=300)
@given(_descripciones)
def test_codigo_falla_igual_a_referencia_literal(g):
    assert codigo_falla_excel(g) == _referencia(g)


@settings(max_examples=300)
@given(st.floats(min_value=-1e6, max_value=1e6, allow_nan=False), st.integers(min_value=0, max_value=4))
def test_redondeo_idempotente_y_cercano(x, d):
    r = redondear_excel(x, d)
    assert redondear_excel(r, d) == r
    assert abs(r - x) <= 0.5 * 10**-d + 1e-9 * max(1.0, abs(x))
