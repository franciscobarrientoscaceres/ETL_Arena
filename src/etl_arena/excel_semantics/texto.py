"""Conversión número→texto de fórmulas Excel y código de falla de ``ListOfFaults!F`` (F-14)."""

from __future__ import annotations


def texto_excel(v: object) -> str:
    """Texto que produce Excel al usar ``v`` en una función de texto (``FIND``, ``CONCATENATE``).

    Números en formato General: hasta 15 dígitos significativos, sin ceros finales,
    enteros sin decimales, exponente como ``1E+16``. Booleanos: ``TRUE``/``FALSE``.
    Separador decimal ``.``: en el libro solo hay fallas enteras, así que la configuración
    regional no interviene.
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        s = f"{float(v):.15g}"
        if "e" in s:
            mantisa, exp = s.split("e")
            signo = "-" if exp.startswith("-") else "+"
            s = f"{mantisa}E{signo}{abs(int(exp)):02d}"
        return s
    return str(v)


def codigo_falla_excel(g: object) -> str:
    """``=IFERROR(MID(G,1,FIND(" ",G,1)-1),CONCATENATE("F",G))``.

    * sin espacio → ``FIND`` da error → ``"F" & G`` (``169`` → ``"F169"``);
    * espacio en la posición 1 → ``MID(G,1,0)`` → ``""`` (no entra al ``IFERROR``);
    * en otro caso → el texto antes del primer espacio (``"F55 EXTERNAL…"`` → ``"F55"``).
    """
    s = texto_excel(g)
    p = s.find(" ")
    return "F" + s if p == -1 else s[:p]
