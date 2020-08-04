from utils.functions import *
import shutil
import os


class GitCommandLines():
    def __init__(self):
        self.repository = r'https://github.com/asahi8769/Consult2.git'
        self.abs_dir = None
        self.rel_dir = None
        subprocess_cmd (f'git config --global user.name Ilhee Lee')
        subprocess_cmd (f'git config --global user.email asahi8769@gmail.com')

    def push_rep(self):
        self.clone_rep()
        self.manage_pulls()
        if self.ask_overwrite() is False:
            os.startfile(self.rel_dir)
            return
        subprocess_cmd (f'git init')
        subprocess_cmd (f'git add .')
        subprocess_cmd (f'git config --global http.sslVerify false')
        subprocess_cmd (f'git commit -m "Apply all changes"')
        subprocess_cmd (f'git remote add origin {self.repository}')
        subprocess_cmd (f'git push --force origin master')
        subprocess_cmd (f'git remote remove origin')

    def clone_rep(self):
        self.abs_dir = make_pulled_dir()
        self.rel_dir = os.path.relpath(self.abs_dir, os.getcwd())
        subprocess_cmd(f'git rm -rf --cached .')
        subprocess_cmd(f'git clone --depth=1 {self.repository[:-4]} {self.rel_dir}')

    def history(self):
        subprocess_cmd (f'git log ')

    def ask_overwrite(self):
        if input('Overwrite current repository? (y/n, Default : n) :').lower() != 'y':
            return False

    def manage_pulls(self):
        if len(sorted(os.listdir('pulled'), reverse=True)) > 3:
            print('Pulls :', len(sorted(os.listdir('pulled'), reverse=True)))
            shutil.rmtree(os.path.join(
                'pulled', sorted(os.listdir('pulled'), reverse=True)[3-len(sorted(os.listdir('pulled'), reverse=True))]),ignore_errors=True)
            # os.rmdir(os.path.join('pulled', sorted(os.listdir('pulled'), reverse=True)[-1]))


if __name__ == "__main__":
    GitCommandLines().push_rep()