import ftplib

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql


def dfo_update:
    try:
        with ftplib.FTP(lp.ftp_ip, lp.ftp_user, lp.ftp_psw) as ftp:
            ftp.cwd('functions')
            with open("./functions/" + lp.ftp_dfoppy, "wb") as f:
                ftp.retrbinary("RETR " + lp.ftp_dfoppy, f.write)
    except:
        source = r"C:\Apache\htdocs\CED_Cagliari\CED\functions\dataframeoperations.py"
        dest = r"C:\Apache\htdocs\CED_Automi\automi\functions\dataframeoperations.py"
        shutil.copy(source, dest)
