"""Sesión COM de Excel propia y aislada (tarea 4.1; ADR-05, design.md §workbook).

* Crea **su propia** instancia (``DispatchEx``), nunca se engancha a un Excel abierto por el usuario.
* Visible (las macros usan ``.Select``/``ActiveWindow``), sin alertas, macros habilitadas solo en esta
  instancia (``AutomationSecurity = msoAutomationSecurityLow``; no cambia la configuración de Office).
* Al salir cierra **solo su instancia**: cierra los libros sin guardar, ``Quit`` y, si el proceso
  sigue vivo (referencias COM colgadas), lo termina por PID. Un watchdog lo mata si excede el timeout.
* Un vigilante cierra los diálogos de error de VBA ("Microsoft Visual Basic", p. ej. error 438) con
  "Finalizar" y guarda su texto: la macro falla con ese mensaje en segundos, no al vencer el timeout.
"""

from __future__ import annotations

import contextlib
import gc
import logging
import threading
import time
from pathlib import Path

log = logging.getLogger("etl_arena.excel")

MSO_AUTOMATION_SECURITY_LOW = 1
STILL_ACTIVE = 259
TITULOS_DIALOGO_VBA = ("Microsoft Visual Basic", "Microsoft Visual Basic for Applications")
BOTONES_FINALIZAR = ("&Finalizar", "&End", "Finalizar", "End")


class ErrorExcel(RuntimeError):
    """Excel no disponible, macro fallida o timeout."""


class ErrorVBA(ErrorExcel):
    """Error en tiempo de ejecución de una macro (texto del diálogo de VBA)."""

    def __init__(self, macro: str, mensaje: str, segundos: float):
        super().__init__(f"{macro}: error VBA: {mensaje}")
        self.macro, self.mensaje, self.segundos = macro, mensaje, segundos


def _proceso_vivo(pid: int) -> bool:
    import win32api
    import win32con
    import win32process

    try:
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    except Exception:
        return False
    try:
        return win32process.GetExitCodeProcess(h) == STILL_ACTIVE
    finally:
        win32api.CloseHandle(h)


def _terminar(pid: int) -> None:
    import win32api
    import win32con

    try:
        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 1)
        win32api.CloseHandle(h)
        log.warning("instancia de Excel %s terminada por PID", pid)
    except Exception:  # ya terminó
        pass


def _cerrar_dialogos_vba(pid: int) -> list[str]:
    """Busca diálogos de error de VBA del proceso, devuelve su texto y los cierra con "Finalizar"."""
    import win32con
    import win32gui
    import win32process

    mensajes: list[str] = []

    def por_ventana(h, _):
        try:
            if (
                win32process.GetWindowThreadProcessId(h)[1] != pid
                or win32gui.GetClassName(h) != "#32770"
                or win32gui.GetWindowText(h) not in TITULOS_DIALOGO_VBA
            ):
                return
        except Exception:  # la ventana se cerró mientras se enumeraba
            return
        textos, boton = [], None

        def por_hijo(c, _):
            nonlocal boton
            texto = win32gui.GetWindowText(c)
            if win32gui.GetClassName(c) == "Static" and texto:
                textos.append(texto)
            elif texto in BOTONES_FINALIZAR:
                boton = c

        win32gui.EnumChildWindows(h, por_hijo, None)
        if boton:
            mensajes.append(" ".join(" ".join(t.split()) for t in textos))
            win32gui.PostMessage(boton, win32con.BM_CLICK, 0, 0)

    win32gui.EnumWindows(por_ventana, None)
    return mensajes


class SesionExcel:
    def __init__(self, visible: bool = True, timeout_s: float = 900):
        self.visible = visible
        self.timeout_s = timeout_s
        self.app = None
        self.pid: int | None = None
        self._libros: list = []
        self._watchdog: threading.Timer | None = None
        self.vencida = False
        self.errores_vba: list[str] = []
        self._vigilando = threading.Event()

    # ------------------------------------------------------------------ ciclo de vida
    def __enter__(self) -> SesionExcel:
        import pythoncom
        import win32com.client
        import win32process

        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            pythoncom.CoUninitialize()
            raise ErrorExcel(f"no se pudo iniciar Excel por COM: {exc}") from exc
        _, self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)
        self.app.Visible = self.visible
        self.app.DisplayAlerts = False
        self.app.AskToUpdateLinks = False
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW  # solo esta instancia
        self._watchdog = threading.Timer(self.timeout_s, self._vencer)
        self._watchdog.daemon = True
        self._watchdog.start()
        self._vigilando.set()
        threading.Thread(target=self._vigilar_dialogos, daemon=True).start()
        self.version = f"{self.app.Version} build {self.app.Build}"
        self.build = int(self.app.Build)
        log.info("Excel %s iniciado (PID %s)", self.version, self.pid)
        return self

    def _vigilar_dialogos(self) -> None:
        while self._vigilando.is_set():
            try:
                mensajes = _cerrar_dialogos_vba(self.pid)
            except Exception:  # una ventana que se cierra a mitad de la enumeración; se reintenta
                log.debug("enumeración de diálogos interrumpida", exc_info=True)
                mensajes = []
            for mensaje in mensajes:
                log.error("diálogo de error VBA cerrado: %s", mensaje)
                self.errores_vba.append(mensaje)
            time.sleep(1)

    def _vencer(self) -> None:
        self.vencida = True
        if self.pid:
            log.error("timeout de %.0f s: se termina Excel (PID %s)", self.timeout_s, self.pid)
            _terminar(self.pid)

    def __exit__(self, *exc) -> None:
        import pythoncom

        if self._watchdog:
            self._watchdog.cancel()
        self._vigilando.clear()
        try:
            # Los proxies de libro se sueltan antes de Quit: liberarlos con Excel ya cerrado da
            # RPC_E_DISCONNECTED (0x80010108). Quien llama tampoco debe retener ``wb`` al salir.
            while self._libros:
                wb = self._libros.pop()
                with contextlib.suppress(Exception):  # el libro ya pudo cerrarse
                    wb.Close(SaveChanges=False)
                del wb
            gc.collect()
            if self.app is not None:
                with contextlib.suppress(Exception):  # si Quit falla, se termina por PID abajo
                    self.app.Quit()
        finally:
            self._libros.clear()
            self.app = None
            gc.collect()
            pythoncom.CoUninitialize()
            if self.pid:
                for _ in range(20):  # hasta ~10 s para que Excel termine solo
                    if not _proceso_vivo(self.pid):
                        break
                    time.sleep(0.5)
                else:
                    _terminar(self.pid)

    # ------------------------------------------------------------------ operaciones
    def abrir(self, ruta: str | Path):
        """Abre un libro para escritura, sin actualizar vínculos."""
        ruta = Path(ruta).resolve()
        if not ruta.exists():
            raise ErrorExcel(f"no existe el libro {ruta}")
        wb = self.app.Workbooks.Open(str(ruta), UpdateLinks=0, ReadOnly=False)
        self._libros.append(wb)
        return wb

    def ejecutar_macro(self, wb, macro: str) -> float:
        """``Application.Run`` de una macro del libro. Devuelve los segundos que tardó.

        Si la macro muestra un error de VBA, el vigilante lo cierra y aquí se lanza ``ErrorVBA``
        con el texto del diálogo."""
        wb.Activate()
        previos = len(self.errores_vba)
        inicio = time.perf_counter()
        try:
            self.app.Run(f"'{wb.Name}'!{macro}")
        except Exception as exc:
            time.sleep(1.5)  # el vigilante puede estar guardando el texto del diálogo
            if self.vencida:
                raise ErrorExcel(f"{macro}: timeout de {self.timeout_s:.0f} s") from exc
            if len(self.errores_vba) > previos:
                raise ErrorVBA(macro, self.errores_vba[-1], time.perf_counter() - inicio) from exc
            raise ErrorExcel(f"{macro} falló: {exc}") from exc
        return time.perf_counter() - inicio
