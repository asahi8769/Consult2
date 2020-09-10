import os
from utils.functions import packaging, subprocess_cmd, path_find


if __name__ == "__main__":
    dir_venv_64 = os.path.join(os.getcwd(), 'venv', 'Scripts')
    dir_loc = os.path.join(os.getcwd(), 'dist')
    # file_to_compile = path_find('Main_Window.py', os.getcwd())
    file_to_compile = path_find('Main_Window.py', os.getcwd())
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = os.path.join(os.getcwd(), 'flying_bird.ico').replace('\\', '/')
    # install_command = f'pyinstaller.exe -F --noconsole --hidden-import=xlrd --icon={icon_name} {file_to_compile}'
    install_command = f'pyinstaller.exe -F --noconsole --hidden-import=xlrd {file_to_compile}'
    # install_command = f'pyinstaller.exe -F --hidden-import=xlrd {file_to_compile}'
    print(install_command)

    subprocess_cmd(f'cd {dir_venv_64} & {install_command} & cd dist & copy Main_Window.exe {dir_loc}')
    # packaging(r'Main_Window.exe', 'Cookies', 'Images', 'Main_DB', 'Objections', 'Spawn', 'Cookies_objection')
    os.startfile('dist')
