import os
import urllib.request
import zipfile
import shutil
import time


manifestation_master_directory = os.getcwd()
manifestation_file = manifestation_master_directory+"\\Manifestation.exe"
Old_manifestation_directory = manifestation_master_directory+"\\manifestation_exe-master"

proxy_handler = urllib.request.ProxyHandler({'https': 'http://proxy-dmz.intel.com:912'})
opener = urllib.request.build_opener(proxy_handler)
urllib.request.install_opener(opener)


def installation():
    urllib.request.urlretrieve("https://github.com/idriss-animashaun-intel/manifestation/archive/refs/heads/master.zip", manifestation_master_directory+"\\manifestation_luancher_new.zip")
    print("*** Updating Launcher Please Wait ***")
    zip_ref = zipfile.ZipFile(manifestation_master_directory+"\\manifestation_luancher_new.zip", 'r')
    zip_ref.extractall(manifestation_master_directory)
    zip_ref.close()
    os.remove(manifestation_master_directory+"\\manifestation_luancher_new.zip")

    src_dir = manifestation_master_directory + "\\manifestation-master"
    dest_dir = manifestation_master_directory
    fn = os.path.join(src_dir, "Manifestation.exe")
    shutil.copy(fn, dest_dir)

    shutil.rmtree(manifestation_master_directory+"\\manifestation-master")

    time.sleep(5)
    
def upgrade():
    print("*** Updating Launcher Please Wait ***")    
    print("*** Removing old files ***")
    time.sleep(20)
    os.remove(manifestation_file)
    time.sleep(10)
    installation()


### Is manifestation already installed? If yes get file size to compare for upgrade
if os.path.isfile(manifestation_file):
    local_file_size = int(os.path.getsize(manifestation_file))
    # print(local_file_size)

    url = 'https://github.com/idriss-animashaun-intel/manifestation/raw/master/Manifestation.exe'
    f = urllib.request.urlopen(url)

    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)

    if local_file_size != web_file_size:# upgrade available
        upgrade()

### manifestation wasn't installed, so we download and install it here                
else:
    installation()

if os.path.isdir(Old_manifestation_directory):
        print('removing manifestation_exe-master')
        time.sleep(5)
        shutil.rmtree(Old_manifestation_directory)

print('Launcher up to date')