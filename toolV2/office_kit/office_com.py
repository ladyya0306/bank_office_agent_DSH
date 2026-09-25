"""Microsoft Office automation over COM, with explicit availability probing.

Using the installed Office 2016 gives real fidelity (native PDF export, format
conversion) that pure-Python libraries cannot match.  Every entry point here
degrades gracefully when Office is missing or a document is locked.
"""
from __future__ import annotations

import contextlib
import threading
from pathlib import Path
from typing import Any, Iterator

from .common import OfficeKitError

# COM apartment state must be set once per thread before any Dispatch call.
_TLS = threading.local()
_AVAILABILITY: dict[str, bool | str] = {}

PROGIDS = {"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"}

WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32
MSO_TRUE = -1
MSO_FALSE = 0

# XlFileFormat constants used by excel_to_format
XL_CSV = 6
XL_XLSX = 51
XL_XLS = 56
XL_HTML = 44


class ComError(OfficeKitError):
    """Office automation is unavailable or the operation failed."""


def _init_com() -> None:
    if getattr(_TLS, "initialized", False):
        return
    try:
        import pythoncom  # type: ignore

        pythoncom.CoInitialize()
        _TLS.pythoncom = pythoncom
    except ImportError:
        pass
    _TLS.initialized = True


def available(app: str) -> bool:
    """Whether a given Office application can actually be started.

    The probe starts a real instance and then quits it. The COM reference must be
    released explicitly - merely dropping the Python name can leave a zombie
    WINWORD.EXE/EXCEL.EXE holding a document lock on Windows.
    """
    if app in _AVAILABILITY:
        return bool(_AVAILABILITY[app])
    _init_com()
    real = None
    excel_before: set[int] = set()
    try:
        import win32com.client as wc

        # DispatchEx starts a private instance instead of attaching to a Word/Excel
        # the user already has open, so Quit() can never close their documents.
        # Snapshot BEFORE dispatch: only processes absent from this set are ours.
        excel_before = _excel_pids() if app == "excel" else set()
        try:
            real = wc.DispatchEx(PROGIDS[app])
        except Exception:  # noqa: BLE001
            real = wc.Dispatch(PROGIDS[app])
        _AVAILABILITY[app] = True
    except Exception as exc:  # noqa: BLE001
        _AVAILABILITY[app] = f"{type(exc).__name__}: {exc}"
    finally:
        if real is not None:
            _force_quit(real, app)
            del real
            _release_com()
        if app == "excel":
            _kill_new_excel(excel_before)
    return bool(_AVAILABILITY[app])


def _force_quit(application, app: str) -> None:
    """Close open documents then quit, so no zombie process is left behind."""
    with contextlib.suppress(Exception):
        if app == "word":
            for _ in range(application.Documents.Count):
                application.Documents(1).Close(False)
        elif app == "excel":
            for _ in range(application.Workbooks.Count):
                application.Workbooks(1).Close(False)
        elif app == "ppt":
            for _ in range(application.Presentations.Count):
                application.Presentations(1).Close()
    with contextlib.suppress(Exception):
        application.Quit()


def _excel_pids(*, poll: bool = False) -> set[int]:
    """PIDs of running EXCEL.EXE processes, for detecting ones we spawn."""
    try:
        import psutil
    except ImportError:
        return set()

    def snapshot() -> set[int]:
        out: set[int] = set()
        for proc in psutil.process_iter(["name", "pid"]):
            try:
                if (proc.info.get("name") or "").upper() == "EXCEL.EXE":
                    out.add(int(proc.info["pid"]))
            except Exception:  # noqa: BLE001
                continue
        return out

    result = snapshot()
    if poll and not result:
        # A fresh COM server may not have appeared in the process table yet.
        import time

        for _ in range(10):
            time.sleep(0.3)
            result = snapshot()
            if result:
                break
    return result


def _kill_new_excel(before: set[int]) -> int:
    """Terminate EXCEL.EXE processes that appeared during our automation.

    Excel is a single-instance COM server and routinely ignores Quit(), leaving
    worker processes behind that keep the workbook file locked (verified: they
    stay alive indefinitely). Only processes that were NOT in ``before`` are
    touched, so an Excel the user already had open is never closed.
    """
    try:
        import psutil
    except ImportError:
        return 0
    killed = 0
    for pid in _excel_pids() - before:
        try:
            proc = psutil.Process(pid)
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except Exception:  # noqa: BLE001
                proc.kill()
                with contextlib.suppress(Exception):
                    proc.wait(timeout=3)
            killed += 1
        except Exception:  # noqa: BLE001
            continue
    return killed


def _release_com() -> None:
    """Drop this thread's COM references so the server can actually exit."""
    pythoncom = getattr(_TLS, "pythoncom", None)
    if pythoncom is None:
        return
    with contextlib.suppress(Exception):
        import gc

        gc.collect()
    with contextlib.suppress(Exception):
        pythoncom.CoFreeUnusedLibraries()


def availability_report() -> dict[str, Any]:
    out: dict[str, Any] = {}
    for app in PROGIDS:
        if available(app):
            out[app] = {"available": True}
        else:
            out[app] = {"available": False, "error": _AVAILABILITY.get(app)}
    return out


@contextlib.contextmanager
def office_app(app: str, *, visible: bool = False) -> Iterator[Any]:
    """Start an Office application and guarantee it is quit afterwards."""
    if not available(app):
        raise ComError(
            f"Microsoft {app} is not reachable via COM ({_AVAILABILITY.get(app)}); "
            f"use the pure-Python fallback instead"
        )
    _init_com()
    import win32com.client as wc

    real = None
    excel_before: set[int] = set()
    try:
        # Snapshot BEFORE dispatch: Excel may already have an instance that we
        # must leave alone, and only processes absent from this set are ours.
        excel_before = _excel_pids() if app == "excel" else set()
        try:
            real = wc.DispatchEx(PROGIDS[app])
        except Exception:  # noqa: BLE001
            real = wc.Dispatch(PROGIDS[app])
        with contextlib.suppress(Exception):
            real.Visible = visible
        with contextlib.suppress(Exception):
            real.DisplayAlerts = 0
        yield real
    except OfficeKitError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise ComError(f"{app} automation failed: {exc}") from exc
    finally:
        if real is not None:
            _force_quit(real, app)
            del real
            _release_com()
        if app == "excel":
            _kill_new_excel(excel_before)


def _apath(path: str | Path) -> str:
    """Office COM always wants a fully qualified path."""
    return str(Path(path).resolve())


# --------------------------------------------------------------------------
# Word
# --------------------------------------------------------------------------
def _with_doc(app_name: str, opener, worker) -> Any:
    """Open one Office document, run ``worker``, always close and release COM.

    Releasing the document reference (and collecting) before the app quits is what
    keeps WINWORD.EXE/EXCEL.EXE from lingering as a zombie holding a file lock.
    """
    with office_app(app_name) as app:
        doc = None
        try:
            doc = opener(app)
            return worker(doc)
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)
                del doc
                _release_com()


def word_to_pdf(src: Path, dst: Path) -> Path:
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    _with_doc(
        "word",
        lambda app: app.Documents.Open(_apath(src), ReadOnly=True, AddToRecentFiles=False),
        lambda doc: doc.ExportAsFixedFormat(_apath(dst), WD_FORMAT_PDF),
    )
    return dst


def word_save_as(src: Path, dst: Path, fmt: int) -> Path:
    """fmt: 2=txt, 6=rtf, 8=html, 12/16/17=docx."""
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    _with_doc(
        "word",
        lambda app: app.Documents.Open(_apath(src), ReadOnly=True, AddToRecentFiles=False),
        lambda doc: doc.SaveAs2(_apath(dst), FileFormat=fmt),
    )
    return dst


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def excel_to_pdf(src: Path, dst: Path) -> Path:
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    _with_doc(
        "excel",
        lambda app: app.Workbooks.Open(_apath(src), ReadOnly=True, UpdateLinks=0),
        lambda wb: wb.ExportAsFixedFormat(XL_TYPE_PDF, _apath(dst)),
    )
    return dst


def excel_to_format(src: Path, dst: Path, fmt: int) -> Path:
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    _with_doc(
        "excel",
        lambda app: app.Workbooks.Open(_apath(src), ReadOnly=True, UpdateLinks=0),
        lambda wb: wb.SaveAs(_apath(dst), FileFormat=fmt),
    )
    return dst


# --------------------------------------------------------------------------
# PowerPoint
# --------------------------------------------------------------------------
def ppt_to_pdf(src: Path, dst: Path) -> Path:
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    _with_doc(
        "ppt",
        lambda app: app.Presentations.Open(_apath(src), WithWindow=MSO_FALSE, ReadOnly=MSO_TRUE),
        lambda pres: pres.SaveAs(_apath(dst), PP_SAVE_AS_PDF),
    )
    return dst


def ppt_to_images(src: Path, out_dir: Path, fmt: str = "PNG") -> list[Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    produced: list[Path] = []

    def export(pres):
        for i, slide in enumerate(pres.Slides, start=1):
            target = out_dir / f"slide_{i:03d}.{fmt.lower()}"
            slide.Export(_apath(target), fmt)
            produced.append(target)

    _with_doc(
        "ppt",
        lambda app: app.Presentations.Open(_apath(src), WithWindow=MSO_FALSE, ReadOnly=MSO_TRUE),
        export,
    )
    return produced


# --------------------------------------------------------------------------
# bulk conversion: one app instance reused across many documents
# --------------------------------------------------------------------------
def word_batch_to_pdf(pairs: list[tuple[Path, Path]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    with office_app("word") as app:
        for src, dst in pairs:
            entry: dict[str, Any] = {"source": str(src), "target": str(dst)}
            doc = None
            try:
                Path(dst).parent.mkdir(parents=True, exist_ok=True)
                doc = app.Documents.Open(_apath(src), ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(_apath(dst), WD_FORMAT_PDF)
                entry["ok"] = True
            except Exception as exc:  # noqa: BLE001
                entry["ok"] = False
                entry["error"] = str(exc)
            finally:
                if doc is not None:
                    with contextlib.suppress(Exception):
                        doc.Close(False)
                    del doc
                    _release_com()
            results.append(entry)
    return results


def excel_batch_to_pdf(pairs: list[tuple[Path, Path]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    with office_app("excel") as app:
        for src, dst in pairs:
            entry: dict[str, Any] = {"source": str(src), "target": str(dst)}
            wb = None
            try:
                Path(dst).parent.mkdir(parents=True, exist_ok=True)
                wb = app.Workbooks.Open(_apath(src), ReadOnly=True, UpdateLinks=0)
                wb.ExportAsFixedFormat(XL_TYPE_PDF, _apath(dst))
                entry["ok"] = True
            except Exception as exc:  # noqa: BLE001
                entry["ok"] = False
                entry["error"] = str(exc)
            finally:
                if wb is not None:
                    with contextlib.suppress(Exception):
                        wb.Close(False)
                    del wb
                    _release_com()
            results.append(entry)
    return results


def ppt_batch_to_pdf(pairs: list[tuple[Path, Path]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    with office_app("ppt") as app:
        for src, dst in pairs:
            entry: dict[str, Any] = {"source": str(src), "target": str(dst)}
            pres = None
            try:
                Path(dst).parent.mkdir(parents=True, exist_ok=True)
                pres = app.Presentations.Open(_apath(src), WithWindow=MSO_FALSE, ReadOnly=MSO_TRUE)
                pres.SaveAs(_apath(dst), PP_SAVE_AS_PDF)
                entry["ok"] = True
            except Exception as exc:  # noqa: BLE001
                entry["ok"] = False
                entry["error"] = str(exc)
            finally:
                if pres is not None:
                    with contextlib.suppress(Exception):
                        pres.Close()
                    del pres
                    _release_com()
            results.append(entry)
    return results


LEGACY_MAP = {".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx"}


def docx_to_legacy_pdf_fallback(src: Path, dst: Path) -> Path:
    """When Office is missing, still give the user *a* PDF (text-level fidelity).

    It re-lays-out the extracted text rather than reproducing Word's pagination,
    so callers must surface this as a lower-fidelity result.
    """
    from .doc_ops import render_text_pdf

    return render_text_pdf(src, dst)
