"""Logging estructurado (design.md §Error Handling): una línea JSON por evento con ``id_corrida``."""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime

_ESTANDAR = set(vars(logging.makeLogRecord({}))) | {"message", "asctime"}


class FormatoJson(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        datos = {
            "ts": datetime.fromtimestamp(record.created).isoformat(timespec="seconds"),
            "nivel": record.levelname,
            "logger": record.name,
            "mensaje": record.getMessage(),
        }
        datos.update({k: v for k, v in vars(record).items() if k not in _ESTANDAR})  # p. ej. id_corrida
        if record.exc_info:
            datos["excepcion"] = self.formatException(record.exc_info)
        return json.dumps(datos, ensure_ascii=False, default=str)


def configurar_logging(json_: bool = True, nivel: int = logging.INFO) -> None:
    manejador = logging.StreamHandler(sys.stderr)
    manejador.setFormatter(FormatoJson() if json_ else logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    raiz = logging.getLogger("etl_arena")
    raiz.handlers[:] = [manejador]
    raiz.setLevel(nivel)
    raiz.propagate = False
