import os
from pathlib import Path


def _decorator(func):
    def wrapper(dirname):
        cwd = os.getcwd()
        if os.path.basename(cwd).startswith('consult') or os.path.basename(cwd).startswith('Consult'):
            pass
        else:
            os.chdir(os.path.join(Path(cwd).parent))
            func(dirname)
            os.chdir(cwd)
    return wrapper


@_decorator
def make_dir(dirname):
    try:
        os.mkdir(dirname)
        print("Directory ", dirname, " Created ")
    except FileExistsError:
        pass


make_dir('Main_DB')
make_dir('Cookies')
make_dir('Images')
make_dir('Spawn')
make_dir('Objections')
make_dir('GUI')
make_dir('utils')
