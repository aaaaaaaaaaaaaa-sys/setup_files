import zipfile
import tempfile
import os,sys
import shutil
import subprocess
from tqdm import tqdm
import urllib.request
import comtypes.client
import pathlib
import uuid
import hashlib
import winreg
import json
import urllib.request
from colorama import init, Fore, Back, Style
import gdown_mine as gdl
from tqdm import tqdm

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)
def progress_print(block_count, block_size, total_size):
    percentage = 100.0 * block_count * block_size / total_size
    if percentage > 100:
        percentage = 100
    max_bar = 50
    bar_num = int(percentage / (100 / max_bar))
    progress_element = '=' * bar_num
    if bar_num != max_bar:
        progress_element += '>'
    bar_fill = ' ' # これで空のとこを埋める
    bar = progress_element.ljust(max_bar, bar_fill)
    total_size_kb = total_size / 1024
    print(
        Fore.LIGHTCYAN_EX,
        f'[{bar}] {percentage:.2f}% ( {total_size_kb:.0f}KB )\r',
        end=''
    )

def download(url, filepath_and_filename):
    init()
    print(Fore.WHITE,"必要なファイルをダウンロードしています")
    urllib.request.urlretrieve(url, filepath_and_filename, progress_print)
    print('') # 改行
    print(Style.RESET_ALL, end="")

def extractprogress(block_count,total_size):
    percentage = int(block_count)
    if percentage > 100:
        percentage = 100
    max_bar = 50
    bar_num = int(percentage / (100 / max_bar))
    progress_element = '=' * bar_num
    if bar_num != max_bar:
        progress_element += '>'
    bar_fill = ' ' 
    bar = progress_element.ljust(max_bar, bar_fill)
    print(
        Fore.LIGHTCYAN_EX,
        f'[{bar}] {percentage:.2f}%\r',
        end=''
    )

def decompress(filename,outpath):
    init()
    print(Fore.WHITE,"ファイルを展開しています...")
    try:
        cmd = [resource_path('./bin/7zr.exe'),"x","-mmt=on",f"-o{outpath}","-bsp1","-y",filename]
        with subprocess.Popen(cmd, stdout=subprocess.PIPE, bufsize=1,universal_newlines=True) as p:
            for line in p.stdout:
                line = line.replace("\n", "")
                if "%" in line:
                    extractprogress(line.split("%")[0],100)
        extractprogress(100,100)
    except:
        pass
    try:
        with zipfile.ZipFile(filename) as zipf:
            bar = tqdm(total=len(zipf.infolist()))
            for info in zipf.infolist():
                info.filename = info.orig_filename.encode('cp437').decode('cp932')
                if os.sep != "/" and os.sep in info.filename:
                    info.filename = info.filename.replace(os.sep, "/")
                bar.update(1)
                bar.write(info.filename)
                zipf.extract(info,outpath)
    except:
        pass
    bar.close()
    print('') # 改行
    print(Style.RESET_ALL, end="")

def yes_no_input(showstring):
    while True:
        choice = input(f"{showstring}[y/n]>>>").lower()
        if choice in ['y', 'ye', 'yes']:
            return True
        elif choice in ['n', 'no']:
            return False
def checkrun(processname):
    cmd = 'tasklist'
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    running = False
    for line in proc.stdout:
        if str(processname) in str(line):
            running = True
    return running
def makelnk(targetpath,savepath,comment=""):
    try:
        os.remove(savepath)
    except :
        pass
    try:
        wsh=comtypes.client.CreateObject("wScript.Shell",dynamic=True)
        short=wsh.CreateShortcut(savepath)
        short.TargetPath=targetpath
        short.Description=comment
        short.workingDirectory=os.path.dirname(targetpath)
        short.Save()
    except :
        pass

def gethash(filepath):
    try:
        algo = 'md5'
        hash = hashlib.new(algo)
        Length = hashlib.new(algo).block_size * 0x800
        with open(filepath,'rb') as f:
            BinaryData = f.read(Length)
            while BinaryData:
                hash.update(BinaryData)
                BinaryData = f.read(Length)
        return hash.hexdigest()
    except :
        return None


runname = "実行するファイル名"
installpath = "展開するフォルダ名"
dlurl = "ダウンロードするファイルのURL"
lnkname = "ショートカットの名前"
menuname = "スタートメニューの名前"
lnkcomment = "ショートカットのコメント"
uninstallname = "コントロールパネルに表示される名前"
uninstallset = False
askrunname = "実行するときに聞く名前"
menuset = False #右クリックメニューに追加するか
rightname = "右クリックメニューで表示される名前"

try:
    if sys.argv[1] == "uninstall":
        localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
        DST = localpath + installpath
        if os.path.exists(DST):
            if yes_no_input("アンインストールしますか?"):
                if checkrun(os.path.basename(runname)):
                    if yes_no_input(f"実行中のプロセス({os.path.basename(runname)})が見つかりました、停止しますか?(停止することをお勧めします)"):
                        try:
                            subprocess.run(["taskkill","/f","/im",os.path.basename(runname)])
                        except :
                            pass
                    else:
                        pass
                else:
                    pass
                if os.path.exists(DST):
                    maxcount = 0
                    for rootdir,subdir,files in os.walk(DST):
                        for file in files:
                            maxcount += 1
                    try:
                        with open(os.path.join(DST,'install_data.json'), mode='rt', encoding='utf-8') as file:
                            installjson = json.load(file)
                    except:
                        pass
                    try:
                        uuiddata = installjson["uninstall_key"]
                        path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,path)
                    except:
                        import traceback
                        traceback.print_exc()
                    progress = tqdm(total=maxcount)
                    for rootdir,subdir,files in os.walk(DST):
                        for file in files:
                            filepath = os.path.join(rootdir,file)
                            progress.update(1)
                            try:
                                os.remove(filepath)
                                progress.write(f"{filepath}---Success")
                            except :
                                progress.write(f"{filepath}---Failed")
                    try:
                        shutil.rmtree(DST)
                    except :
                        pass
                    progress.close()
                    if os.name == 'nt':
                        home = os.getenv('USERPROFILE')
                    else:
                        home = os.getenv('HOME')
                    desktop_dir = os.path.join(home, 'Desktop')
                    menudir = f"C:\\Users\\{os.getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\{menuname}"
                    try:
                        shutil.rmtree(menudir)
                    except :
                        pass
                    try:
                        os.remove(os.path.join(desktop_dir,lnkname))
                    except:
                        pass
                    fileid = str(uuid.uuid4())
                    with open(os.path.join(os.environ["TMP"],str(fileid)) + ".bat","w",encoding="shift-jis") as tempmake:
                        tempmake.write(f'@echo off\nSET dir={DST}\n:loop\nif exist %dir% goto CODE1\nif not exist %dir% goto CODE2\n:CODE1\nrd /s /q {DST}\ngoto :loop\n:CODE2\ndel /f "%~dp0%~nx0"')
                    try:
                        subprocess.Popen([os.path.abspath(os.path.join(os.environ["TMP"],str(fileid)) + ".bat")])            
                    except:
                        import traceback
                        traceback.print_exc()
                    uninstallset = True
except:
    pass
if uninstallset:
    sys.exit(1)
while True:
    selectlist = ["1","2","3"]
    localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
    DST = localpath + installpath
    if os.path.exists(DST):
        print("１．アップデートを確認する")
    else:
        print("１．新規でインストールする")
    print("２．再インストールする(使用履歴は残ります)")
    print("３．アンインストールする")
    print("閉じる場合は×ボタンで閉じてください")
    print()
    userselect = input("実行する操作番号を入力してください(半角)(1,2,3)>>>")
    if str(userselect) in selectlist:
        if str(userselect) == "1":
            if os.path.exists(DST):
                with tempfile.TemporaryDirectory() as tmpdir:
                    downloadpath = os.path.join(tmpdir,'donwloaded_file.zip')
                    try:
                        os.makedirs(os.path.dirname(downloadpath))
                    except :
                        pass
                    localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
                    DST = localpath + installpath
                    url = dlurl
                    download(url,downloadpath)
                    if gethash(f"{DST}\sample.zip") == gethash(downloadpath):
                        print("変更はありません")
                    else:
                        if yes_no_input("変更が見つかりました、アップデートしますか?"):
                            if checkrun(os.path.basename(runname)):
                                if yes_no_input(f"実行中のプロセス({os.path.basename(runname)})が見つかりました、停止しますか?(停止することをお勧めします)"):
                                    try:
                                        subprocess.run(["taskkill","/f","/im",os.path.basename(runname)])
                                    except :
                                        pass
                                else:
                                    pass
                            decompress(downloadpath,DST)
                            try:
                                samplepath = f"{DST}\\sample.zip"
                                try:
                                    os.remove(samplepath)
                                except :
                                    pass
                            except :
                                pass
                            shutil.copyfile(downloadpath,samplepath)
                            if os.name == 'nt':
                                home = os.getenv('USERPROFILE')
                            else:
                                home = os.getenv('HOME')
                            desktop_dir = os.path.join(home, 'Desktop')
                            menudir = f"C:\\Users\\{os.getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\{menuname}"
                            try:
                                os.makedirs(menudir)
                            except :
                                pass
                            if menuset:
                                try:
                                    path = f'Software\\Classes\\*\\shell\\{rightname}\\command'
                                    key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                                    winreg.SetValueEx(key, '', 0, winreg.REG_SZ,str('''"{}" "%V"'''.format(DST + "\\" + runname)))
                                    winreg.CloseKey(key)
                                except:
                                    import traceback
                                    traceback.print_exc()
                            uuiddata = str(uuid.uuid4())
                            data = {
                                "installpath":str(DST),
                                "uninstall_key":uuiddata,
                                "rightkey":str(path)}
                            with open(os.path.join(DST,'install_data.json'), mode='wt', encoding='utf-8') as file:
                                json.dump(data, file, ensure_ascii=False, indent=2)
                            makelnk(DST + "\\" + runname,os.path.join(desktop_dir,lnkname),comment=lnkcomment)
                            makelnk(DST + "\\" + runname,os.path.join(menudir,lnkname),comment=lnkcomment)
                            path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                            try:
                                key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                                winreg.SetValueEx(key, 'DisplayName', 0, winreg.REG_SZ,str(uninstallname))
                                winreg.SetValueEx(key, 'UninstallString', 1, winreg.REG_SZ,str(os.path.abspath(sys.argv[0])) + " uninstall")
                                winreg.CloseKey(key)
                            except:
                                import traceback
                                traceback.print_exc()
                            if yes_no_input(f"{askrunname}を実行しますか?"):
                                try:
                                    subprocess.Popen(["start",os.path.abspath(DST + "\\" + runname)])
                                except :
                                    pass
                                print("実行できない場合はスタートメニュー等から実行してください")
            else:
                with tempfile.TemporaryDirectory() as tmpdir:
                    downloadpath = os.path.join(tmpdir,'donwloaded_file.zip')
                    localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
                    DST = localpath + installpath
                    if os.path.exists(DST):
                        if checkrun(os.path.basename(runname)):
                            if yes_no_input(f"実行中のプロセス({os.path.basename(runname)})が見つかりました、停止しますか?(停止することをお勧めします)"):
                                try:
                                    subprocess.run(["taskkill","/f","/im",os.path.basename(runname)])
                                except :
                                    pass
                            else:
                                pass
                        if yes_no_input("フォルダが存在します、上書きしますか?"):
                            try:
                                shutil.rmtree(DST)
                            except :
                                pass
                            try:
                                os.makedirs(DST)
                            except :
                                pass
                        else:
                            continue
                    else:
                        try:
                            os.makedirs(DST)
                        except :
                            pass
                    url = dlurl
                    download(url,downloadpath)
                    decompress(downloadpath,DST)
                    try:
                        samplepath = f"{DST}\\sample.zip"
                        try:
                            os.remove(samplepath)
                        except :
                            pass
                        shutil.copyfile(downloadpath,samplepath)
                    except :
                        pass
                    if os.name == 'nt':
                        home = os.getenv('USERPROFILE')
                    else:
                        home = os.getenv('HOME')
                    desktop_dir = os.path.join(home, 'Desktop')
                    menudir = f"C:\\Users\\{os.getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\{menuname}"
                    try:
                        os.makedirs(menudir)
                    except :
                        pass
                    path = "None"
                    if menuset:
                        try:
                            path = f'Software\\Classes\\*\\shell\\{rightname}\\command'
                            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                            winreg.SetValueEx(key, '', 0, winreg.REG_SZ,str('''"{}" "%V"'''.format(DST + "\\" + runname)))
                            winreg.CloseKey(key)
                        except:
                            import traceback
                            traceback.print_exc()
                    uuiddata = str(uuid.uuid4())
                    data = {
                        "installpath":str(DST),
                        "uninstall_key":uuiddata,
                        "rightkey":str(path)}
                    with open(os.path.join(DST,'install_data.json'), mode='wt', encoding='utf-8') as file:
                        json.dump(data, file, ensure_ascii=False, indent=2)
                    makelnk(DST + "\\" + runname,os.path.join(desktop_dir,lnkname),comment=lnkcomment)
                    makelnk(DST + "\\" + runname,os.path.join(menudir,lnkname),comment=lnkcomment)
                    try:
                        try:
                            os.remove(os.path.join(DST,"setup.exe"))
                        except:
                            pass
                        shutil.copyfile(os.path.abspath(sys.argv[0]),os.path.join(DST,"setup.exe"))
                        try:
                            path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                            winreg.SetValueEx(key, 'DisplayName', 0, winreg.REG_SZ,str(uninstallname))
                            winreg.SetValueEx(key, 'UninstallString', 1, winreg.REG_SZ,str(os.path.join(DST,"setup.exe")) + " uninstall")
                            winreg.CloseKey(key)
                        except:
                            import traceback
                            traceback.print_exc()
                    except:
                        pass
                    makelnk(DST + "\\" + runname,os.path.join(desktop_dir,lnkname),comment=lnkcomment)
                    makelnk(DST + "\\" + runname,os.path.join(menudir,lnkname),comment=lnkcomment)
                    if yes_no_input(f"{askrunname}を実行しますか?"):
                        try:
                            subprocess.Popen(["start",os.path.abspath(DST + "\\" + runname)])
                        except :
                            pass
                        print("実行できない場合はスタートメニュー等から実行してください")
                    break
        elif str(userselect) == "2":
            if yes_no_input("再インストールしますか?"):
                localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
                DST = localpath + installpath
                if checkrun(runname):
                    if yes_no_input(f"実行中のプロセス({runname})が見つかりました、停止しますか?(停止することをお勧めします)"):
                        try:
                            subprocess.run(["taskkill","/f","/im",runname])
                        except :
                            pass
                    else:
                        pass
                else:
                    pass
                with tempfile.TemporaryDirectory() as tmpdir:
                    downloadpath = os.path.join(tmpdir,'donwloaded_file.zip')
                    try:
                        os.makedirs(os.path.dirname(downloadpath))
                    except :
                        pass
                    url = dlurl
                    download(url,downloadpath)
                    decompress(downloadpath,DST)
                    try:
                        samplepath = f"{DST}\\sample.zip"
                        try:
                            os.remove(samplepath)
                        except :
                            pass
                        shutil.copyfile(downloadpath,samplepath)
                    except :
                        pass
                if os.name == 'nt':
                    home = os.getenv('USERPROFILE')
                else:
                    home = os.getenv('HOME')
                desktop_dir = os.path.join(home, 'Desktop')
                menudir = f"C:\\Users\\{os.getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\{menuname}"
                try:
                    os.makedirs(menudir)
                except :
                    pass
                try:
                    try:
                        os.remove(os.path.join(DST,"setup.exe"))
                    except:
                        pass
                    shutil.copyfile(os.path.abspath(sys.argv[0]),os.path.join(DST,"setup.exe"))
                    try:
                        with open(os.path.join(DST,'install_data.json'), mode='rt', encoding='utf-8') as file:
                            installjson = json.load(file)
                    except:
                        pass
                    try:
                        uuiddata = installjson["uninstall_key"]
                        path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,path)
                    except:
                        import traceback
                        traceback.print_exc()
                    try:
                        path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                        winreg.SetValueEx(key, 'DisplayName', 0, winreg.REG_SZ,str(uninstallname))
                        winreg.SetValueEx(key, 'UninstallString', 1, winreg.REG_SZ,str(os.path.join(DST,"setup.exe")) + " uninstall")
                        winreg.CloseKey(key)
                    except:
                        import traceback
                        traceback.print_exc()
                except:
                    pass
                makelnk(DST + "\\" + runname,os.path.join(desktop_dir,lnkname),comment=lnkcomment)
                makelnk(DST + "\\" + runname,os.path.join(menudir,lnkname),comment=lnkcomment)
                if yes_no_input(f"{askrunname}を実行しますか?"):
                    try:
                        subprocess.Popen(["start",os.path.abspath(DST + "\\" + runname)])
                    except :
                        pass
                    print("実行できない場合はスタートメニュー等から実行してください")
                break
        elif str(userselect) == "3":
            localpath = os.environ.get("LOCALAPPDATA") + "\\Programs"
            DST = localpath + installpath
            if os.path.exists(DST):
                if yes_no_input("アンインストールしますか?"):
                    if checkrun(os.path.basename(runname)):
                        if yes_no_input(f"実行中のプロセス({os.path.basename(runname)})が見つかりました、停止しますか?(停止することをお勧めします)"):
                            try:
                                subprocess.run(["taskkill","/f","/im",os.path.basename(runname)])
                            except :
                                pass
                        else:
                            pass
                    else:
                        pass
                    if os.path.exists(DST):
                        maxcount = 0
                        for rootdir,subdir,files in os.walk(DST):
                            for file in files:
                                maxcount += 1
                        try:
                            with open(os.path.join(DST,'install_data.json'), mode='rt', encoding='utf-8') as file:
                                installjson = json.load(file)
                        except:
                            pass
                        try:
                            uuiddata = installjson["uninstall_key"]
                            path = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{uuiddata}"
                            winreg.DeleteKey(winreg.HKEY_CURRENT_USER,path)
                        except:
                            import traceback
                            traceback.print_exc()
                        progress = tqdm(total=maxcount)
                        for rootdir,subdir,files in os.walk(DST):
                            for file in files:
                                filepath = os.path.join(rootdir,file)
                                progress.update(1)
                                try:
                                    os.remove(filepath)
                                    progress.write(f"{filepath}---Success")
                                except :
                                    progress.write(f"{filepath}---Failed")
                        try:
                            shutil.rmtree(DST)
                        except :
                            pass
                        progress.close()
                        if os.name == 'nt':
                            home = os.getenv('USERPROFILE')
                        else:
                            home = os.getenv('HOME')
                        if menuset:
                            path = f'Software\\Classes\\*\\shell'
                            try:
                                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path + "\\{rightname}\\command")
                                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path + "\\{rightname}")
                            except:
                                pass
                        desktop_dir = os.path.join(home, 'Desktop')
                        menudir = f"C:\\Users\\{os.getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\{menuname}"
                        try:
                            shutil.rmtree(menudir)
                        except :
                            pass
                        try:
                            os.remove(os.path.join(desktop_dir,lnkname))
                        except:
                            pass
                        fileid = str(uuid.uuid4())
                        with open(os.path.join(os.environ["TMP"],str(fileid)) + ".bat","w",encoding="shift-jis") as tempmake:
                            tempmake.write(f'@echo off\nSET dir={DST}\n:loop\nif exist %dir% goto CODE1\nif not exist %dir% goto CODE2\n:CODE1\nrd /s /q {DST}\ngoto :loop\n:CODE2\ndel /f "%~dp0%~nx0"')
                        try:
                            subprocess.Popen(["start",os.path.abspath(os.path.join(os.environ["TMP"],str(fileid)) + ".bat")])            
                        except:
                            pass
                        break    
                    else:
                        print("フォルダが存在しません")
                        continue
                else:
                    continue
            else:
                print("ディレクトリが存在しませんでした")
                continue
    else:
        print("無効な数値が入力されました")
        continue