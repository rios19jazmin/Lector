from cx_Freeze import setup, Executable

setup(
    name="lector1.1",
    version="1.1",
    description="lector de codigos",
    executables=[Executable("main.py")]
)
