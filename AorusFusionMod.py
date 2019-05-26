# 
#   https://github.com/Shamilius/AorusFusionMod
#
import ctypes, sys, os, win32process, win32gui, win32con, win32api, subprocess, xml.etree.ElementTree as ET
from time import sleep

#print(subprocess.STD_OUTPUT_HANDLE)

#sleep(1)

#os.startfile(r'C:\Program Files (x86)\AorusFusion\switchProfile1.exe')
#os.system('"'+Profiles[SetProfile][0]+'"')
#subprocess.call(r'C:\Program Files (x86)\AorusFusion\switchProfile1.exe', shell=True)

#sys.exit()

#   Check if program runs with administrator rights
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

print('Program runs with administrator rights:', bool(is_admin()))

#   Run the program with administrator rights if compiled
if not is_admin() and sys.argv[0][sys.argv[0].__len__()-3:] != '.py': 
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

#   Remove useless Vista_EQ_C.exe which runs with profile change
if os.path.exists('Vista_EQ_C.exe'): os.remove('Vista_EQ_C.exe')


AorusFusionPath = os.environ['PROGRAMFILES']+'\\AorusFusion\\'
ProfileDataXML = AorusFusionPath+'ProfileData.xml'
ActiveWindow = ''
LastWindow = ''
CurrentProfile = -1


#
#   LoadProfiles() return example: [{
#       '1': ['switchProfile1.exe', ''],
#       '2': ['switchProfile2.exe', 'xrEngine.exe'],
#       '2': ['switchProfile3.exe', 'javaw.exe']
#   }, 1558615973]
#
def LoadProfiles():
    Profiles = dict()

    #   Read ProfileData.xml
    try:
        ProfilesXML = ET.parse(ProfileDataXML).getroot()
    except Exception as e:
        print('Exception reading '+ProfileDataXML+':', e)
        sys.exit()

    for Profile in ProfilesXML:
        for ProfileData in Profile:
            if ProfileData.tag != 'Name': continue
            if ProfileData.text == 'null': continue

            App = ProfileData.text.lower()
            Profiles[Profile.attrib['ID']] = ['switchProfile'+Profile.attrib['ID']+'.exe', (App if App.find('.exe') > -1 else '')]

    print('Profiles:', Profiles)
    return [Profiles, os.stat(ProfileDataXML)[8]]


#   Profiles: dictionary with profiles data
#   ProfilesDataModified: Last modified timestamp - for readable date/time use time.ctime(timestamp)
Profiles, ProfilesDataModified = LoadProfiles()

while True:
    #   See what is current ActiveWindow
    hwnd = win32gui.GetForegroundWindow()
    if hwnd == 0:
        sleep(0.5)
        continue

    try:
        hndl = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, 0, win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception as e:
        sleep(1)
        print('Exception:', e)
        #   Maybe AorusFusion is opened - check if ProfileData.xml was changed
        if ProfilesDataModified != os.stat(ProfileDataXML)[8]: Profiles, ProfilesDataModified = LoadProfiles()
        continue

    ActiveWindowPath = win32process.GetModuleFileNameEx(hndl, 0)
    ActiveWindow = ActiveWindowPath[ActiveWindowPath.rfind('\\')+1:].lower()

    #   AorusFusion is opened - check if ProfileData.xml was changed
    if ActiveWindow == 'aorusfusion.exe':
        sleep(1)
        #   AorusFusion is opened - check if ProfileData.xml was changed
        if ProfilesDataModified != os.stat(ProfileDataXML)[8]: Profiles, ProfilesDataModified = LoadProfiles()
        continue

    #   if ActiveWindow didn't changed
    if ActiveWindow == LastWindow:
        sleep(0.5)
        continue

    print('ActiveWindow:', ActiveWindow)

    #   if ActiveWindow was changed
    for Profile in Profiles:
        if Profiles[Profile][1] == ActiveWindow:
            SetProfile = Profile
            break
        elif Profiles[Profile][1] == '':
            SetProfile = Profile
    
    if CurrentProfile != SetProfile:
        subprocess.Popen([Profiles[SetProfile][0]])
        #os.startfile(Profiles[SetProfile][0])
        CurrentProfile = SetProfile

    LastWindow = ActiveWindow
    sleep(0.5)

