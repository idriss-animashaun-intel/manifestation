import os
import urllib.request
import zipfile
import shutil
import time
from subprocess import Popen
# from subprocess import call


manifest_tool_master_directory = os.getcwd()
manifest_tool_directory = manifest_tool_master_directory+"\manifestation-updates"
manifest_tool_file = manifest_tool_directory+"\\main.exe"

manifest_tool_old_directory = manifest_tool_master_directory+"\\old_revisions"
manifest_tool_old_file = manifest_tool_old_directory+"\\main.exe"

proxy_handler = urllib.request.ProxyHandler({'https': 'http://proxy-dmz.intel.com:912'})
opener = urllib.request.build_opener(proxy_handler)
urllib.request.install_opener(opener)

def installation():
    print("*** Downloading new version ***")
    urllib.request.urlretrieve("https://github.com/idriss-animashaun-intel/manifestation/archive/refs/heads/updates.zip", manifest_tool_master_directory+"\\manifest_tool_new.zip")
    print("*** Extracting new version ***")
    zip_ref = zipfile.ZipFile(manifest_tool_master_directory+"\manifest_tool_new.zip", 'r')
    zip_ref.extractall(manifest_tool_master_directory)
    zip_ref.close()
    os.remove(manifest_tool_master_directory+"\manifest_tool_new.zip")
    time.sleep(5)
    
def upgrade():    
    print("*** Removing old files ***")
    shutil.rmtree(manifest_tool_directory)
    time.sleep(10)
    installation()


### Is manifest_tool already installed? If yes get file size to compare for upgrade
if os.path.isfile(manifest_tool_file):
    local_file_size = int(os.path.getsize(manifest_tool_file))
    # print(local_file_size)
    ### Check if update needed:
    f = urllib.request.urlopen("https://github.com/idriss-animashaun-intel/manifestation/raw/updates/main.exe") # points to the exe file for size
    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)

    if local_file_size != web_file_size:# upgrade available
        updt = input("*** New upgrade available! enter <y> to upgrade now, other key to skip upgrade *** ")
        if updt == "y": # proceed to upgrade
            upgrade()

### For the transfer between GitHub and GitLab
elif os.path.isfile(manifest_tool_old_file):
    installation()

### manifest_tool wasn't installed, so we download and install it here                
else:
    install = input("Welcome to manifest_tool! If you enter <y> manifest_tool will be downloaded in the same folder where this file is.\nAfter the installation, this same file you are running now (\"manifest_tool.exe\") will the one to use to open manifest_tool :)\nEnter any other key to skip the download\n -->")
    if install == "y":
        installation()

print('Ready')


### We open the real application:
try:
    Popen(manifest_tool_file)
    print("*** Opening Manifest Summary Tool ***")
    time.sleep(20)
except:
    print('Failed to open application, Please open manually in subfolder')
    pass