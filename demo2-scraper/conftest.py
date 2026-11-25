"""Fija el directorio del demo en sys.path para importar scraper desde los tests."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
