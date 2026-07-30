"""Extension repair: restore stripped video extensions via magic bytes, or
bulk-convert one extension to another. Emits Ops for the journaled executor."""

import os
from pathlib import Path

from .parse import VIDEO_EXTS, is_junk_dir, is_junk_name
from .plan import Op, Plan, check_collisions

_KNOWN_NON_VIDEO = {'.nfo', '.srt', '.sub', '.idx', '.jpg', '.jpeg', '.png',
                    '.txt', '.nzb', '.bat', '.sh', '.py', '.json', '.xlsx',
                    '.xls', '.jsonl',
                    # Our own crash scar: without this, a restore pass would
                    # magic-byte sniff it and append a video extension.
                    '.mediaorg_tmp'}


def detect_extension(filepath: Path) -> str | None:
    """Read the file header to determine the video format."""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(12)
    except OSError:
        return None
    if header.startswith(b'\x1a\x45\xdf\xa3'):
        return '.mkv'
    if len(header) >= 8 and header.startswith(b'ftyp', 4):
        return '.mp4'
    if header.startswith(b'RIFF') and header.startswith(b'AVI ', 8):
        return '.avi'
    if header.startswith(b'\x00\x00\x01'):
        return '.mpg'
    if header.startswith(b'\x30\x26\xb2\x75'):
        return '.wmv'
    if header.startswith(b'FLV\x01'):
        return '.flv'
    return None


def plan_extension_restore(folder: Path) -> Plan:
    """Find files missing a video extension and plan restoring it."""
    ops: list[Op] = []
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames[:] = sorted(d for d in dirnames if not is_junk_dir(d))
        for filename in sorted(f for f in filenames if not is_junk_name(f)):
            ext = Path(filename).suffix.lower()
            if ext in VIDEO_EXTS or ext in _KNOWN_NON_VIDEO:
                continue
            filepath = Path(dirpath) / filename
            detected = detect_extension(filepath)
            if detected:
                ops.append(Op("move", filepath,
                              filepath.with_name(filename + detected)))
    return check_collisions(ops)


def plan_extension_convert(folder: Path, from_ext: str, to_ext: str) -> Plan:
    """Plan renaming every *.from_ext under folder to *.to_ext."""
    from_ext = '.' + from_ext.lstrip('.').lower()
    to_ext = '.' + to_ext.lstrip('.').lower()
    ops: list[Op] = []
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames[:] = sorted(d for d in dirnames if not is_junk_dir(d))
        for filename in sorted(f for f in filenames if not is_junk_name(f)):
            p = Path(dirpath) / filename
            if p.suffix.lower() == from_ext:
                ops.append(Op("move", p, p.with_suffix(to_ext)))
    return check_collisions(ops)
