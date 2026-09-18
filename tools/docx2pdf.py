#!/usr/bin/env python3
"""Converte .docx em .pdf pela ponte UNO do LibreOffice, atualizando
sumário e campos (paginação) antes da exportação."""

import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

PROFILE = os.environ.get(
    "LO_PROFILE", "/tmp/claude-lo-profile"
)
PIPE = "ccpipe"


def _prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _start_soffice():
    os.makedirs(PROFILE, exist_ok=True)
    return subprocess.Popen(
        [
            "soffice",
            "-env:UserInstallation=file://" + PROFILE,
            "--headless",
            "--invisible",
            "--nologo",
            "--norestore",
            "--nolockcheck",
            "--accept=pipe,name=%s;urp;" % PIPE,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _connect(timeout=120):
    ctx_local = uno.getComponentContext()
    resolver = ctx_local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", ctx_local
    )
    url = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % PIPE
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        try:
            return resolver.resolve(url)
        except Exception as exc:  # aguarda o listener subir
            last = exc
            time.sleep(1)
    raise RuntimeError("nao foi possivel conectar ao LibreOffice: %s" % last)


def convert(src, dst):
    proc = _start_soffice()
    try:
        ctx = _connect()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(src)),
            "_blank",
            0,
            (_prop("Hidden", True), _prop("ReadOnly", False)),
        )
        # varias passagens: atualizar o sumario altera a paginacao, que
        # precisa ser refletida de volta nos campos de pagina.
        indexes = doc.getDocumentIndexes()
        for _ in range(3):
            doc.getTextFields().refresh()
            doc.refresh()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
        doc.getTextFields().refresh()
        doc.refresh()
        doc.storeToURL(
            uno.systemPathToFileUrl(os.path.abspath(dst)),
            (_prop("FilterName", "writer_pdf_Export"),),
        )
        doc.close(False)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        time.sleep(2)
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                proc.kill()


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("uso: docx2pdf.py <entrada.docx> <saida.pdf>")
    convert(sys.argv[1], sys.argv[2])
    print("PDF gerado: %s" % sys.argv[2])
