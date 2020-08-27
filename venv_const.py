from utils.functions import subprocess_cmd
import os

dir_venv_64 = os.path.join(os.getcwd(), 'venv', 'Scripts')


def install(lib):
    return f'pip --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org install {lib}'


# subprocess_cmd(f'cd {dir_venv_64} & '
#                # f'{install("tqdm")} & '
#                # f'{install("openpyxl")} & '
#                # f'{install("matplotlib")} & '
#                # f'{install("pyinstaller")} & '
#                # f'{install("xlrd")} & '
#                # f'{install("pandas")} & '
#                # f'{install("pyautogui")}'
#                f'{install("pulp")} & '
#                )

subprocess_cmd(f'cd {dir_venv_64} & {install("pulp")} & {install("scipy")}')

