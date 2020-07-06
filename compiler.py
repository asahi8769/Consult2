import os
from utils.functions import packaging, subprocess_cmd, path_find


if __name__ == "__main__":
    dir_py_32 = r'C:\Users\glovis-laptop\AppData\Local\Programs\Python\Python37-32\Scripts'.replace('\\', '/')
    dir_py_64 = r'C:\Users\glovis-laptop\AppData\Local\Programs\Python\Python37\Scripts'.replace('\\', '/')
    dir_venv_64 = r'D:\devs\consult2\venv\Scripts'
    dir_loc = os.path.join(os.getcwd(), 'dist')
    file_to_compile = path_find('Main_Window.py', r'D:\devs\consult2')
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = r'D:\devs\consult2\flying_bird.ico'.replace('\\', '/')

    # install_command = f'pyinstaller.exe -F --noconsole --hidden-import=matplotlib.pyplot --hidden-import=sqlite3 --hidden-import=pandas --hidden-import=tqdm --hidden-import=smtplib --hidden-import=email.mime.image --hidden-import=email.mime.text --hidden-import=email.mime.multipart --hidden-import=email.mime.application --icon={icon_name} {file_to_compile}'
    install_command = f'pyinstaller.exe -F --noconsole --hidden-import=xlrd --icon={icon_name} {file_to_compile}'
    print(install_command)

    subprocess_cmd(f'cd {dir_venv_64} & {install_command} & cd dist & copy Main_Window.exe {dir_loc}')
    packaging(r'Main_Window.exe', 'Cookies', 'Images', 'Main_DB', 'Objections', 'Spawn', 'Cookies_objection')
