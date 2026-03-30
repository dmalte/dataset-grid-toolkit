from pathlib import Path
import tempfile
import os


def atomic_write_text(path: Path, text: str):
    """Atomically write text to a file (write to temp, then rename)."""
    dirp = path.parent
    dirp.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(dirp))
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(text)
        os.replace(tmp, str(path))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass
