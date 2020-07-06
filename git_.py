from utils.functions import subprocess_cmd, old_ver_directory
import os, shutil



class GitCommandLines():
    def __init__(self):
        self.repository = r'https://github.com/asahi8769/Warranty_Consult2.git'
        subprocess_cmd (f'git config --global user.name Ilhee Lee')
        subprocess_cmd(f'git config --global user.email asahi8769@gmail.com')

    def push_rep(self):
        self.clone_rep()
        subprocess_cmd (f'git init')
        subprocess_cmd (f'git add .')
        subprocess_cmd (f'git config --global http.sslVerify false')
        # subprocess_cmd (f'git rm -r --cached images.zip')
        # subprocess_cmd (f'git rm -r --cached .idea/')
        # subprocess_cmd (f'git rm -r --cached __pycache__/')
        subprocess_cmd (f'git commit -m "Add all my files"')
        subprocess_cmd (f'git remote add origin {self.repository}')
        subprocess_cmd (f'git push --force origin master')
        subprocess_cmd(f'git remote remove origin')

    def clone_rep(self):
        # abs_dir =
        rel_dir = os.path.relpath(old_ver_directory(), os.getcwd())
        subprocess_cmd(f'git clone {self.repository[:-4]} {rel_dir}')
        # shutil.rmtree(os.path.join(abs_dir, '.git'))  # deleting .git folder to save space

    def history(self):
        subprocess_cmd(f'git log ')


if __name__ == "__main__":
    GitCommandLines().push_rep()