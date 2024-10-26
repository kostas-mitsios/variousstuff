from cx_Freeze import setup, Executable

setup(
    name="boss_music",
    version="1.0",
    description="Your application description",
    executables=[Executable("boss_music.py")]
)
