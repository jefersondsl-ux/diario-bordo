from PyInstaller.utils.hooks import collect_all

streamlit_datas, streamlit_binaries, streamlit_hiddenimports = collect_all("streamlit")

a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=streamlit_binaries,
    datas=[
        ("diario_bordo_entrada.py", "."),
        *streamlit_datas,
    ],
    hiddenimports=[
        *streamlit_hiddenimports,
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "pandas",
    ],
    hookspath=[],
    noarchive=False,
    
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DiarioBordo",
    debug=False,
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name="DiarioBordo",
)