# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for Xplor Data Migration Tools
# =====================================================
#
# BUILD INSTRUCTIONS
# ------------------
# Run these commands once to install build tools:
#   Windows:  py -m pip install pyinstaller streamlit pandas openpyxl xlrd lxml
#   Mac:      pip3 install pyinstaller streamlit pandas openpyxl xlrd lxml
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
# - Expected output size: ~450–650 MB (Streamlit + Pandas are large packages)
# - Expected startup time: 5–15 seconds (Streamlit starts a local web server)
# ─────────────────────────────────────────────────────────────────────────────

import sys
from PyInstaller.utils.hooks import collect_all

# ── Collect dependencies ──────────────────────────────────────────────────────
st_datas,     st_binaries,     st_hiddenimports     = collect_all("streamlit")
pd_datas,     pd_binaries,     pd_hiddenimports     = collect_all("pandas")
altair_datas, altair_binaries, altair_hiddenimports = collect_all("altair")
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all("openpyxl")
xlrd_datas,   xlrd_binaries,   xlrd_hiddenimports   = collect_all("xlrd")
lxml_datas,   lxml_binaries,   lxml_hiddenimports   = collect_all("lxml")

# ── Analysis ─────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=(
        st_binaries
        + pd_binaries
        + altair_binaries
        + openpyxl_binaries
        + xlrd_binaries
        + lxml_binaries
    ),
    datas=[
        # ── App entry point & core engine ─────────────────────────────────
        ("Home.py",         "."),
        ("validator_v2.py", "."),
        # ── Multi-page app structure ───────────────────────────────────────
        ("pages",           "pages"),    # 1_Xplor_Validator … 5_Balance_Adjustments
        ("shared",          "shared"),   # service_map.py, styles.py
        ("scripts",         "scripts"),  # check_names, prepare_bookings_import, etc.
        ("assets",          "assets"),   # Balance Adjustments template XLSX
    ] + st_datas + pd_datas + altair_datas + openpyxl_datas + xlrd_datas + lxml_datas,
    hiddenimports=(
        st_hiddenimports
        + pd_hiddenimports
        + altair_hiddenimports
        + openpyxl_hiddenimports
        + xlrd_hiddenimports
        + lxml_hiddenimports
        + [
            # Streamlit internals
            "streamlit.runtime.scriptrunner.magic_funcs",
            "streamlit.web.cli",
            "streamlit.components.v1",
            # openpyxl styling modules (dynamically loaded)
            "openpyxl.styles",
            "openpyxl.utils",
            "openpyxl.writer.excel",
            # lxml — used by pandas.read_html for HTML-disguised XLS files
            "lxml.etree",
            "lxml.html",
            # xlrd — used for true binary .xls files
            "xlrd",
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
            # tkinter — used by _pick_folder() on Windows (native folder dialog)
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
        bundle_identifier="com.xplor.datamigration.tools",
        info_plist={
            "CFBundleName":               "Xplor Migration Tools",
            "CFBundleDisplayName":        "Xplor Data Migration Tools",
            "CFBundleShortVersionString": "3.0.0",
            "CFBundleVersion":            "3.0.0",
            "NSHighResolutionCapable":    True,
            "NSHumanReadableCopyright":   "Xplor Technologies",
            # Allow the browser to open localhost URLs from the app
            "NSAppTransportSecurity": {
                "NSAllowsLocalNetworking": True,
            },
        },
    )
