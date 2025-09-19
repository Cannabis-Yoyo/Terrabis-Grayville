# Auto-loaded by Python at startup if present on sys.path
import sys, importlib

try:
    import distutils  # pragma: no cover
except Exception:
    try:
        # Map distutils imports to setuptools' backport
        m = importlib.import_module("setuptools._distutils")
        sys.modules["distutils"] = m
        sys.modules["distutils.version"] = importlib.import_module("setuptools._distutils.version")
    except Exception:
        # Last resort: provide the one symbol UC uses
        from packaging.version import Version
        class LooseVersion(str):
            def __init__(self, v): self.v = Version(str(v))
            def __lt__(self, other): return self.v < Version(str(other))
            def __le__(self, other): return self.v <= Version(str(other))
            def __gt__(self, other): return self.v > Version(str(other))
            def __ge__(self, other): return self.v >= Version(str(other))
            def __eq__(self, other): return self.v == Version(str(other))
        mod = type(__import__("types"))("distutils.version")
        mod.LooseVersion = LooseVersion
        sys.modules["distutils"] = type(mod)("distutils")
        sys.modules["distutils.version"] = mod


