from utils.functions import subprocess_cmd

dir_venv_64 = r'D:\devs\consult2\venv\Scripts'

def install(lib):
    return f'pip --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org install {lib}'


subprocess_cmd(f'cd {dir_venv_64} & {install("tqdm")} & {install("openpyxl")} & {install("matplotlib")} & {install("pyinstaller")} & {install("xlrd")} & {install("pandas")} & {install("pyautogui")}')
