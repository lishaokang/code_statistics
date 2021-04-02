import os
import sys
import platform
import subprocess
import xlwt
import re
from collections import defaultdict

excl_name = "code_statistics.xlsx"
baseDir = r"."
projects_list = ["ProjectName"]


def git_path():
    if platform.system() == "Windows":
        git = r"C:\Program Files\Git\bin\git.exe"
    else:
        git = "git"

    return git


def show_git_revision(project_path):
    git = git_path()

    cmd = "\"{0}\" -C {1} log --format=\"%aI %h %s [%an]\" -1".format(git, project_path)
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    (output, _) = p.communicate()
    cmd_output(output)


def cmd_output(content):
    # print(content.decode('utf-8').encode('gbk', 'backslashreplace'))
    print(content)
    sys.stdout.flush()


def git_fetch(project_path, branch):
    git = git_path()

    index_lock = os.path.join(project_path, ".git", "index.lock")
    if os.path.exists(index_lock):
        os.remove(index_lock)

    os.system("\"{0}\" -C {1} fetch --quiet".format(git, project_path))
    os.system("\"{0}\" -C {1}  checkout --quiet -f  {2}".format(git, project_path, branch))
    os.system("\"{0}\" -C {1} clean --quiet -fd".format(git, project_path))
    os.system("\"{0}\" -C {1} reset --quiet --hard origin/{2}".format(git, project_path, branch))


def git_update(projects_name):
    for projectName in projects_name:
        git_pro_path = baseDir + "\\" + projectName + "\\"
        print(">>> Update project {0} {1}".format(projectName, git_pro_path))
        # show_git_revision(git_pro_path)
        git_fetch(git_pro_path, branch="develop")


def code_analysis(projects_list):
    git = git_path()
    pro_coll = dict()  # project coll
    for project in projects_list:
        git_pro_path = baseDir + "\\" + project + "\\"
        cmd = "\"{0}\" -C {1} shortlog  -sne".format(git, git_pro_path)
        output, _ = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).communicate()
        # get author mail
        user_lines = dict()  # user map lines
        for author in list(set(output.decode("utf-8").splitlines())):
            author_mail = re.split(r'[<>]', author)[1]
            # get author code lines
            total_lines = 0
            cmd = "\"{0}\" -C {1} log --no-merges --author={2} --pretty=tformat: --numstat -- . :(exclude)node_modules :(exclude)checker-service/webapp".format(
                git, git_pro_path, author_mail)
            output, _ = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).communicate()
            code_counts = output.decode("utf-8")
            for codeCount in code_counts.splitlines():
                a = re.split(r'[\s]', codeCount)
                if a[0].isdigit() and a[1].isdigit():
                    total_lines += int(a[0]) + int(a[1])
            user_lines[author_mail] = total_lines
        pro_coll[project] = user_lines
    return pro_coll


def make_excl(projects_list, excl_name):
    pro_coll = code_analysis(projects_list)
    wb = xlwt.Workbook()
    sh1 = wb.add_sheet(excl_name.split(".")[0], cell_overwrite_ok=True)
    sh1.write(0, 0, "Name/Project")
    row = 1
    urow = 1
    cols = []
    rows = []
    userinfo = defaultdict(lambda: 0)
    for pro, uselin in pro_coll.items():
        sh1.write(0, row, pro)
        for mai, lin in uselin.items():
            if mai in cols:
                sh1.write(cols.index(mai) + 1, row, lin)
            else:
                cols.append(mai)
                sh1.write(urow, 0, mai)
                sh1.write(urow, row, lin)
                urow += 1
            userinfo[mai] += lin
        row += 1
        rows.append(pro)
    sh1.write(0, row, "Total")
    sh1.write(urow, 0, "ProdTotal")
    for key, values in userinfo.items():
        sh1.write(cols.index(key) + 1, row, values)
    tt = 0
    for key, values in pro_coll.items():
        tx = 0
        for mai, lin in values.items():
            tx += lin
        sh1.write(urow, rows.index(key) + 1, tx)
        tt += tx
    sh1.write(urow, row, tt)
    wb.save(excl_name)


def main(projects_list):
    distinct = list(set(projects_list))
    # git_update(distinct)
    # exclName = input("Please enter the name of the save file: ")
    make_excl(distinct, excl_name)


if __name__ == '__main__':
    main(projects_list)
