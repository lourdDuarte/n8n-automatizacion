"""Word (.docx) mínimo, sin dependencias: títulos, párrafos con **negrita**, viñetas y separadores.

Se arma a mano (zip + XML) para no depender de python-docx en la PC ni en el servidor.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/>
<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-AR"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Titulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Titulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Nota"><w:name w:val="Nota"/><w:basedOn w:val="Normal"/>
<w:rPr><w:i/><w:color w:val="7F7F7F"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
</w:styles>"""


def _runs(texto: str, extra: str = '') -> str:
    """'**negrita**' -> run en negrita; '[[completar]]' -> run resaltado en amarillo."""
    salida = []
    for parte in re.split(r'(\*\*.+?\*\*|\[\[.+?\]\])', texto):
        if not parte:
            continue
        props = extra
        if parte.startswith('**') and parte.endswith('**'):
            parte, props = parte[2:-2], props + '<w:b/>'
        elif parte.startswith('[[') and parte.endswith(']]'):
            parte, props = parte[2:-2], props + '<w:highlight w:val="yellow"/>'
        rpr = f'<w:rPr>{props}</w:rPr>' if props else ''
        salida.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escape(parte)}</w:t></w:r>')
    return ''.join(salida)


class Documento:
    def __init__(self):
        self._partes: list[str] = []

    def titulo(self, texto: str, nivel: int = 1):
        self._partes.append(f'<w:p><w:pPr><w:pStyle w:val="Titulo{nivel}"/></w:pPr>{_runs(texto)}</w:p>')

    def parrafo(self, texto: str, estilo: str | None = None):
        ppr = f'<w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>' if estilo else ''
        self._partes.append(f'<w:p>{ppr}{_runs(texto)}</w:p>')

    def vineta(self, texto: str):
        self._partes.append('<w:p><w:pPr><w:ind w:left="567" w:hanging="283"/><w:jc w:val="left"/></w:pPr>'
                            + _runs('•\t' + texto) + '</w:p>')

    def separador(self):
        self._partes.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/>'
                            '</w:pBdr></w:pPr></w:p>')

    def salto_pagina(self):
        self._partes.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def guardar(self, ruta: Path):
        cuerpo = ''.join(self._partes)
        documento = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
                     f'{cuerpo}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                     '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
                     '</w:sectPr></w:body></w:document>')
        ruta.parent.mkdir(parents=True, exist_ok=True)
        temporal = ruta.with_suffix('.tmp')
        with zipfile.ZipFile(temporal, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', _CONTENT_TYPES)
            z.writestr('_rels/.rels', _RELS)
            z.writestr('word/_rels/document.xml.rels', _DOC_RELS)
            z.writestr('word/styles.xml', _STYLES)
            z.writestr('word/document.xml', documento)
        temporal.replace(ruta)
        return ruta
