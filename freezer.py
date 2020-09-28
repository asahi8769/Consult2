import os
from utils.functions import packaging, subprocess_cmd, path_find
from utils.config import SCRIPTS_DIR


if __name__ == "__main__":
    file_name_py = 'Main_Window.py'
    file_name_exe = 'Main_Window.exe'
    file_to_compile = path_find(file_name_py, os.getcwd())
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = 'flying_bird.ico'

    """flag snippet: --hidden-import=xlrd --icon={icon_name}, --clear """
    freeze_command_with_icon = f'pyinstaller.exe --onefile --noconsole --hidden-import=xlrd --icon={icon_name} {file_to_compile}'
    freeze_command_without_icon = f'pyinstaller.exe --onefile --noconsole --hidden-import=xlrd {file_to_compile}'
    dist_dir = os.path.join(os.getcwd(), 'dist')
    # install_command = f'pyinstaller.exe -F --noconsole --hidden-import=xlrd --hidden-import=PyQt5 --hidden-import=matplotlib {file_to_compile}'

    subprocess_cmd(f'cd {SCRIPTS_DIR} & {freeze_command_without_icon} & cd dist & '
                   f'copy Main_Window.exe {dist_dir} & copy {file_name_exe} {os.getcwd()}')
    packaging(file_name_exe, 'Cookies', 'Images', 'Main_DB', 'Objections', 'Spawn', 'Cookies_objection')
    os.startfile('dist')
