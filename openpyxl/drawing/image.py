# Copyright (c) 2010-2024 openpyxl

from io import BytesIO

try:
    from PIL import Image as PILImage
    PILImageOk = True
except ImportError:
    PILImageOk = False


class Image:
    """Image in a spreadsheet"""

    _id = 1
    _path = "/xl/media/image{0}.{1}"
    anchor = "A1"
    format = "PNG"

    # Also know as Alt Text, but the xml tag refers to 'descr'
    desc = None

    def __init__(self, img: str | PILImage.Image | BytesIO, desc = None):
        """
        Load the given image in memory.
        """
        assert PILImageOk

        self.ref = img

        if isinstance(img, (str, BytesIO)):
            img = PILImage.open(img)
        assert isinstance(img, PILImage.Image)

        self.width, self.height = img.size
        self.desc = desc
        
        self.format = (getattr(img, 'format', None) or '').upper()

        fp = BytesIO()
        if self.format in ['GIF', 'JPEG', 'PNG', "WMF", "EMF"]:
            if getattr(img, 'fp', None):
                try:
                    # Preserve raw binary
                    img.fp.seek(0)
                    fp = BytesIO(img.fp.read())
                except (AttributeError, ValueError, OSError):
                    img.save(fp, format=self.format)
            else:
                img.save(fp, format=self.format)
        elif self.format:
            # Convert to png
            img.save(fp, format='PNG')
            self.format = 'PNG'
            
        fp.seek(0)
        self._raw_data = fp.getvalue()
        fp.close()
        img.close()


    def _data(self):
        """
        Return image data, convert to supported types if necessary
        """
        return self._raw_data


    @property
    def path(self):
        assert self.format
        return self._path.format(self._id, self.format.lower())


    def __eq__(self, other):
        return self.ref == other.ref


    def _write(self, archive):
        archive.writestr(self.path[1:], self._data())
