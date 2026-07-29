"""Start the current parser source without importing a cached app module."""
from pathlib import Path

source = Path(__file__).with_name("app.py")
exec(compile(source.read_text(encoding="utf-8"), str(source), "exec"), {"__name__": "__main__", "__file__": str(source)})
