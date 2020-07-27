import os
from subprocess import Popen, PIPE
from datetime import datetime
from os.path import basename
from zipfile import ZipFile


def make_dir(dirname):
    try:
        os.mkdir(dirname)
        print("Directory ", dirname, " Created ")
    except FileExistsError:
        pass


def path_find(name, *paths):
    for path in paths:
        for root, dirs, files in os.walk(path):
            if name in files:
                return os.path.join(root, name)


def subprocess_cmd(command):
    print(command)
    try :
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=True)
        proc_stdout = process.communicate()[0].strip()
    except Exception as e:
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=False)
        proc_stdout = process.communicate()[0].strip()
    print(proc_stdout)


def old_ver_directory():
    try:
        os.mkdir(os.path.join(os.getcwd(), 'old'))
    except Exception as e:
        pass
    dir = os.path.join(os.getcwd(), 'old', f'old_ver_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
    os.mkdir(dir)
    return dir


def packaging(filename, *bindings):
    zipname = r'dist\consult2.zip'
    while True:
        if os.path.exists(os.path.join('dist',filename)):
            with ZipFile(zipname, 'w') as zipObj:
                zipObj.write(os.path.join('dist', filename), basename(filename))
                for binding in bindings:
                    for folderName, subfolders, filenames in os.walk(binding):
                        for filename in filenames:
                            filePath = os.path.join(folderName, filename)
                            print(os.path.join(binding, basename(filePath)))
                            zipObj.write(filePath, os.path.join(basename(folderName), basename(filePath)))
            print(f'패키징을 완료하였습니다. {zipname}')
            break
        else:
            print('파일이 존재하지 않습니다.')


def make_pulled_dir():
    try:
        os.mkdir(os.path.join(os.getcwd(), 'pulled'))
    except Exception as e:
        pass
    dir = os.path.join(os.getcwd(), 'pulled', f'pulled_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
    os.mkdir(dir)
    return dir