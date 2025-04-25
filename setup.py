from cx_Freeze import setup, Executable

setup(
    name="Validador de Planilhas",
    version="1.0",
    description="Validação de planilhas com interface Tkinter",
    executables=[Executable("validador_planilha.py", base="Win32GUI")]
)
