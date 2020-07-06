import zipfile
import shutil
import os


def open_zipfile(zip_, filename):
    z = zipfile.ZipFile(zip_, "r")
    with z.open(filename, 'r') as file:
        temp_file = os.path.join(os.environ['temp'], filename)
        temp = open(temp_file, "wb")
        shutil.copyfileobj(file, temp)
        temp.close()
    return temp_file


if __name__ == "__main__":
    from pathlib import Path

    print(Path(os.getcwd ()).parent)
    zip_ = os.path.join(Path(os.getcwd ()).parent, 'Cookies\Pocket.zip')
    print(zip_)
    temp_file = open_zipfile (zip_, 'MainDB_init.txt')
    with open(temp_file, 'rt', encoding='UTF8') as file:
        query = ''.join([line.strip() for line in file]).split(';')[1]
        print(query)

