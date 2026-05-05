# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for Xplor Data Migration Validator
# =========================================================
#
# BUILD INSTRUCTIONS
# ------------------
# Run these commands once to install build tools:
#   Windows:  py -m pip install pyinstaller streamlit pandas openpyxl
#   Mac:      pip3 install pyinstaller streamlit pandas openpyxl
#
# Then build (run from this project folder):
#   Windows:  py -m PyInstaller xplor_validator.spec
#   Mac:      python3 -m PyInstaller xplor_validator.spec
#
# NOTE: You must build separately on each platform.
#       A Windows machine produces the .exe; a Mac produces the .app.
#       Cross-compilation is not supported by PyInstaller.
#
# OUTPUT
# ------
#   Windows: dist\XplorValidator\XplorValidator.exe  (share the whole folder)
#   Mac:     dist\XplorValidator.app                 (share the .app file)
#
# NOTES
# -----
# - Output is a FOLDER on Windows (--onedir), not a single .exe.
#   Zip the entire dist\XplorValidator\ folder for distribution.
# - On Mac, users may need to right-click > Open on the first launch
#   to bypass Gatekeeper, as the app is unsigned.
# - Expected output size: ~400–600 MB (Streamlit + Pandas are large packages)
# - Expected startup time: 5–15 seconds (Streamlit starts a local web server)
# ─────────────────────────────────────────────────────────────────────────────

import sys
from PyInstaller.utils.hooks import collect_all

# ── Collect dependencies ──────────────────────────────────────────────────────
st_datas,     st_binaries,     st_hiddenimports     = collect_all("streamlit")
pd_datas,     pd_binaries,     pd_hiddenimports     = collect_all("pandas")
altair_datas, altair_binaries, altair_hiddenimports = collect_all("altair")
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all("openpyxl")

# ── Analysis ─────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=(
        st_binaries + pd_binaries + altair_binaries + openpyxl_binaries
    ),
    datas=[
        ("app.py",          "."),
        ("validator_v2.py", "."),
    ] + st_datas + pd_datas + altair_datas + openpyxl_datas,
    hiddenimports=(
        st_hiddenimports
        + pd_hiddenimports
        + altair_hiddenimports
        + openpyxl_hiddenimports
        + [
            # Streamlit internals
            "streamlit.runtime.scriptrunner.magic_funcs",
            "streamlit.web.cli",
            "streamlit.components.v1",
            # openpyxl styling modules (dynamically loaded)
            "openpyxl.styles",
            "openpyxl.utils",
            "openpyxl.writer.excel",
            # Data / serialisation
            "pyarrow",
            "pyarrow.vendored.version",
            "pydeck",
            "jsonschema",
            "jsonschema.validators",
            "jsonschema._validators",
            "jsonschema._format",
            "referencing",
            "referencing._core",
            # Stdlib
            "email.mime.multipart",
            "email.mime.text",
            "email.mime.base",
            "pkg_resources.py2_warn",
            # tkinter (folder picker dialog)
            "tkinter",
            "tkinter.filedialog",
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "scipy",
        "sklearn",
        "PIL",
        "cv2",
        "tensorflow",
        "torch",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

# ── Executable ───────────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="XplorValidator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # Windows: show console so users can see startup messages and errors.
    # Mac: console window doesn't apply — the app opens in the browser.
    console=sys.platform == "win32",
    icon=None,  # Provide a .ico (Windows) or .icns (Mac) path here if desired
)

# ── Collect (folder bundle) ───────────────────────────────────────────────────
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="XplorValidator",
)

# ── Mac .app bundle ───────────────────────────────────────────────────────────
# PyInstaller ignores BUNDLE on Windows — safe to leave in for both platforms.
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="XplorValidator.app",
        icon=None,  # Path to a .icns file for the dock icon
        bundle_identifier="com.xplor.datamigration.validator",
        info_plist={
            "CFBundleName":               "Xplor Validator",
            "CFBundleDisplayName":        "Xplor Data Migration Validator",
            "CFBundleShortVersionString": "2.0.0",
            "CFBundleVersion":            "2.0.0",
            "NSHighResolutionCapable":    True,
            "NSHumanReadableCopyright":   "Xplor Technologies",
            # Allow the browser to open localhost URLs from the app
            "NSAppTransportSecurity": {
                "NSAllowsLocalNetworking": True,
            },
        },
    )
