import xml.etree.ElementTree as ET
from pandas import DataFrame
from pandas import ExcelWriter
from pandas import to_numeric
import os
from tkinter import Tk
from tkinter import Button
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import OptionMenu
from tkinter import StringVar
from tkinter import W
from tkinter import Label
import webbrowser
import Update_Launcher


def find_manifest():
    global mani_mid
    ulinc = ET.parse(filename).getroot()
    temp = filename.split("/",20)
    index = temp.index('recipe') + 2
    start_path = ''
    for i in range(0, index):
        start_path += temp[i] + '\\'
    start_path = start_path[:-1]
    for group in ulinc:
        if group.tag == 'ProcessSteps':
            for item in group:
                if item.tag == 'ProcessStep':
                    for recipe in item:
                        if "HANDLER" in recipe.get('type'):
                            end_path = recipe.get('directPath').replace('..\..\..\..\..','')
                            end_path = end_path.replace('..\..\..\..','')
                            mani_mid = r'' + start_path + end_path
                            mani_path = ET.parse(mani_mid).getroot()
                            for sub_group in mani_path:
                                if sub_group.tag == 'ComponentRecipe':
                                    if "SDTC_Recipe_Parameter" in sub_group.get('type'):
                                        for group in sub_group:
                                            return group.text


def get_summary():
    location = loc_mani.get()

    if location == "Manifest":
        manifest = find_manifest()
    else:
        manifest = filename
    man_root = ET.parse(manifest).getroot()

    Parameters = ['TemperatureStacticSetPoint',
    'Imod_Y',
    
    'PIDSystemControl_Eta',
    'TemperatureControlMode',
    'PVSource',
    'TemperatureDynamicEnable',]
    
    exclude = '_for_Non-Cooling'
    exclude1 = 'SourceTH'
    exclude2 = ['_Y','_B1','_Window']
    block=['Thermal Control']
    sp = []
    imody = []
    p = []
    i_val = []
    d = []
    eta = []
    tcm = []
    pvs = []
    pvs_th1 = []
    pvs_th2 = []
    tde = []
    pid_pc = []
    pid_ic = []
    pid_dc = []

    list_all = [sp,imody,eta,tcm,pvs,tde]
    for i in range(0,len(Parameters)):
        parameter= Parameters[i]
        for group in man_root:
            if group.get('name') in block:
                for item in group:
                    if parameter in item.get('name'):
                        if exclude not in item.get('name'):
                            if exclude1 not in item.get('name'):
                                list_all[i].append(item.get('value'))

    Parameters = ['PIDSystemControl_P',
    'PIDSystemControl_I',
    'PIDSystemControl_D',
    'PIDIvac_Pc',
    'PIDIvac_Ic',
    'PIDIvac_Dc']

    exclude3 = '_for_Non-Cooling'

    list_all2 = [p,i_val,d,pid_pc,pid_ic,pid_dc]
    for i in range(0,len(Parameters)):
        parameter= Parameters[i]
        for group in man_root:
            if group.get('name') in block:
                for item in group:
                    if parameter in item.get('name'):
                        if exclude3 not in item.get('name'):
                            list_all2[i].append(item.get('value'))

    Parameters = ['PVSourceTH1', 'PVSourceTH2']

    list_all1 = [pvs_th1,pvs_th2]
    for i in range(0,len(Parameters)):
        parameter= Parameters[i]
        for group in man_root:
            if group.get('name') in block:
                for item in group:
                    if parameter in item.get('name'):
                        if exclude not in item.get('name'):
                            list_all1[i].append(item.get('value'))

    parameter = 'TemperatureStacticSetPoint'
    cs = []
    for group in man_root:
        if group.get('name') in block:
            for item in group:
                if parameter in item.get('name'):
                    cs.append(item.get('name')[0:4])


    imod = []
    
    parameter= 'Imod'
    for group in man_root:
        if group.get('name') in block:
            for item in group:
                if parameter in item.get('name'):
                    if exclude2[0] not in item.get('name'):
                        if exclude2[1] not in item.get('name'):
                            if exclude2[2] not in item.get('name'):
                                imod.append(item.get('value'))




    df = DataFrame(cs, columns =['Control Set'])
    df['Setpoint']=sp
    df['Setpoint']=to_numeric(df['Setpoint'])
    df['TemperatureDynamicEnable']=tde
    df['Imod'] = imod
    df['Imod_Y']=imody
    df['Imod_Y']=to_numeric(df['Imod_Y'])
    df['P']=p
    df['P']=to_numeric(df['P'])
    df['I']=i_val
    df['I']=to_numeric(df['I'])
    df['D']=d
    df['D']=to_numeric(df['D'])
    df['Eta']=eta
    df['Eta']=to_numeric(df['Eta'])
    df['TemperatureControlMode'] = tcm
    df['PVSource'] = pvs
    df['PVSource TH1'] = pvs_th1
    df['PVSource TH2'] = pvs_th2
    df['PIDIvac_Pc']=pid_pc
    df['PIDIvac_Pc']=to_numeric(df['PIDIvac_Pc'])
    df['PIDIvac_Ic']=pid_ic
    df['PIDIvac_Ic']=to_numeric(df['PIDIvac_Ic'])
    df['PIDIvac_Dc']=pid_dc
    df['PIDIvac_Dc']=to_numeric(df['PIDIvac_Dc'])

    df['Control Set'] = df['Control Set'].str.replace('_', '')

    if location == "Manifest":
        paths = [filename.replace('/','\\'), mani_mid, manifest]
    else:
        paths = ['none', 'none', manifest.replace('/','\\')]
    
    path_names = ['Manifest', 'Handler', 'Thermal Recipe']
  
    df1 = DataFrame(path_names, columns =['File'])
    df1['Path']= paths

    
    writer = ExcelWriter('Manifest Summary.xlsx', engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Manifest Summary', index=False)
    df1.to_excel(writer, sheet_name='File Path', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Manifest Summary']

    # Add some cell formats.

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', None)
    worksheet.set_column('F:F', None)
    worksheet.set_column('G:G', None)
    worksheet.set_column('H:H', None)
    worksheet.set_column('I:I', None)
    worksheet.set_column('J:J', 25)
    worksheet.set_column('K:K', 15)
    worksheet.set_column('L:L', 18)
    worksheet.set_column('M:M', 18)
    worksheet.set_column('N:N', 18)
    worksheet.set_column('O:O', 18)
    worksheet.set_column('P:P', 18)

    worksheet1 = writer.sheets['File Path']

    worksheet1.set_column('A:A', 15)
    worksheet1.set_column('B:B', 200)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    writer.close() #added to allow time for file lock to be released
    os.system(r'"Manifest Summary.xlsx"')

def select_file():
    global filename
    filetypes = (
        ('text files', '*.xml'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Selected File',
        message=filename
    )


### Main Root
root = Tk()
root.title('Manifest Summary v1.03')


mainframe = ttk.Frame(root, padding="60 50 60 50")
mainframe.grid(column=0, row=0, sticky=('news'))
mainframe.columnconfigure(0, weight=3)
mainframe.rowconfigure(0, weight=3)

def callback(url):
    webbrowser.open_new(url)

link1 = Label(mainframe, text="Wiki: https://goto/manifestation", fg="blue", cursor="hand2")
link1.grid(row = 0,column = 0, sticky=W, columnspan = 2)
link1.bind("<Button-1>", lambda e: callback("https://gitlab.devtools.intel.com/ianimash/manifest-summary-toolkit/-/wikis/Manifest-Summary-Tool"))

link2 = Label(mainframe, text="IT support contact: idriss.animashaun@intel.com", fg="blue", cursor="hand2")
link2.grid(row = 1,column = 0, sticky=W, columnspan = 2)
link2.bind("<Button-1>", lambda e: callback("https://outlook.com"))

open_button = Button(
    mainframe,
    text='Select Manifest',
    command=select_file,
    bg = 'blue', fg = 'white', font = '-family "SF Espresso Shack" -size 12'
)

open_button.grid(row = 2, column = 0)

button_0 = Button(mainframe, text="Pull Manifest Summary", height = 1, width = 20, command = get_summary, bg = 'green', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_0.grid(row = 3, column = 0, rowspan = 2 )

loc_mani = StringVar(mainframe)
loc_mani.set("Manifest") # default value

sel_prod = OptionMenu(mainframe, loc_mani, "Manifest", "Thermal Recipe")
sel_prod.grid(row = 2, column = 1, sticky=W)

### Main loop
root.mainloop()