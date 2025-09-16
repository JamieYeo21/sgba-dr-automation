from __future__ import absolute_import, division, print_function
import docx
from docx.image.exceptions import UnrecognizedImageError
from docx.image.constants import MIME_TYPE
from docx.image.image import BaseImageHeader
import xml.etree.ElementTree as ET

def _ImageHeaderFactory(stream):
    from docx.image import SIGNATURES

    def read_64(stream):
        stream.seek(0)
        return stream.read(64)

    header = read_64(stream)
    for cls, offset, signature_bytes in SIGNATURES:
        end = offset + len(signature_bytes)
        found_bytes = header[offset:end]
        if found_bytes == signature_bytes:
            return cls.from_stream(stream)
    raise UnrecognizedImageError

class Svg(BaseImageHeader):
    @classmethod
    def from_stream(cls, stream):
        px_width, px_height = cls._dimensions_from_stream(stream)
        return cls(px_width, px_height, 72, 72)

    @property
    def content_type(self):
        return MIME_TYPE.SVG

    @property
    def default_ext(self):
        return "svg"

    @classmethod
    def _dimensions_from_stream(cls, stream):
        stream.seek(0)
        data = stream.read()
        root = ET.fromstring(data)
        width_str = root.attrib["width"]
        height_str = root.attrib["height"]
        width = int(''.join([c for c in width_str if c.isdigit()]))
        height = int(''.join([c for c in height_str if c.isdigit()]))
        return width, height

docx.image.Svg = Svg
docx.image.constants.MIME_TYPE.SVG = 'image/svg+xml'
docx.image.SIGNATURES = tuple(list(docx.image.SIGNATURES) + [(Svg, 0, b'<?xml version=')])
docx.image.image._ImageHeaderFactory = _ImageHeaderFactory