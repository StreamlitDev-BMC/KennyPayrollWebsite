# Lightweight imghdr shim for environments that don't include the stdlib imghdr module.
# Provides a minimal implementation of imghdr.what(filename, h=None)
# Uses Pillow if available for better detection, falls back to header-signature checks.

import io

try:
    from PIL import Image
except Exception:
    Image = None

def what(filename, h=None):
    """
    Minimal drop-in replacement for imghdr.what().
    - filename: path string or file-like or None
    - h: optional header bytes
    Returns image type string like 'jpeg', 'png', 'gif', 'bmp', 'webp', or None.
    """
    try:
        # Normalize header bytes
        if h is None:
            # If filename is bytes-like, treat it as header bytes
            if isinstance(filename, (bytes, bytearray, memoryview)):
                h = bytes(filename)
            # If filename is a file-like object, read head
            elif hasattr(filename, "read"):
                try:
                    pos = filename.tell()
                except Exception:
                    pos = None
                h = filename.read(32)
                try:
                    if pos is not None:
                        filename.seek(pos)
                except Exception:
                    pass
            # If filename is a path, read head from file
            elif isinstance(filename, str):
                try:
                    with open(filename, "rb") as f:
                        h = f.read(32)
                except Exception:
                    h = None
            else:
                return None
        else:
            if isinstance(h, memoryview):
                h = bytes(h)
            elif isinstance(h, bytearray):
                h = bytes(h)

        if not h:
            return None

        # Prefer Pillow for robust detection if available
        if Image is not None:
            try:
                img = Image.open(io.BytesIO(h))
                fmt = img.format
                if fmt:
                    return fmt.lower()
            except Exception:
                # Pillow couldn't determine from the header bytes alone - continue to signature checks
                pass

        b = h
        # Signature-based detection (covers common formats)
        if b.startswith(b'\xff\xd8'):
            return 'jpeg'
        if b.startswith(b'\x89PNG\r\n\x1a\n'):
            return 'png'
        if b[:6] in (b'GIF87a', b'GIF89a'):
            return 'gif'
        if b.startswith(b'BM'):
            return 'bmp'
        if len(b) >= 12 and b[:4] == b'RIFF' and b[8:12] == b'WEBP':
            return 'webp'
        if b.startswith(b'\x00\x00\x01\x00') or b.startswith(b'\x00\x00\x02\x00'):
            return 'ico'
        # Add more signatures if you need them
        return None
    except Exception:
        return None
