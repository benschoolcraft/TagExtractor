import win32com.client
import re
import csv
from tkinter import *
from tkinter.ttk import *


ModuleDefinitions = {
    '1769-IQ16': 'PLCIO_5E769',
    '1769-OB16': 'PLCIO_5EB55',
    '5069-IB16': 'PLCIO_5B719',
    '5069-OB16': 'PLCIO_5B2C5',
    'J-Block Inputs': 'PLCIO_59454',
    'J-Block Outputs': 'PLCIO_5980A'
    }

def run(): 

    ofile  = open('tagout.csv', "w", newline='')
    writer = csv.writer(ofile)

    acad = win32com.client.Dispatch("AutoCAD.Application")

    doc = acad.ActiveDocument   # Document object

    Module = [{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},
              {},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}]
    Components = [{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},
                  {},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}]

    wires = []
    devices = [0] * 32
    descriptions = [''] * 32
    alias = [''] * 32
    ports = 16
    j=0

    aliasPrefix = E_aliasPrefix.get()
    slot = C_slot.get()
    cardType = C_type.get()
    indexPrefix = E_indexPrefix.get()
    indexPostfix = E_indexPostfix.get()


    for entity in acad.ActiveDocument.ModelSpace:

        name = entity.EntityName

        if name == 'AcDbBlockReference':

            HasAttributes = entity.HasAttributes

            EffectiveName = entity.EffectiveName

            if (HasAttributes & (EffectiveName == ModuleDefinitions[C_pickModule.get()])):
                foundFirst = False
                offset = 0
                for attrib in entity.GetAttributes():
                    i=0
                    found = False
                    while not found:
                        if i < 9:
                            searchstr = '\w+0' + str(i+ 1)
                        else:
                            searchstr = '\w+' + str(i + 1)
                        search = re.search(searchstr, attrib.TagString)
                        if search:
                            found = True
                            if not foundFirst:
                                offset = i
                                foundFirst = True
                            Module[i-offset][attrib.TagString] = attrib.TextString
                        else:
                            i+=1
                        if i > 64:
                            found = True
                        

            else:
                searchTerminal = re.search('HT0W01', entity.EffectiveName)
                searchPJ = re.search('HC01PJ_1', entity.EffectiveName)
                searchDest = re.search('HA1D', entity.EffectiveName)
                searchWD = re.match('WD', entity.EffectiveName)
                if (searchTerminal == None) & (searchDest == None) & (searchWD == None) & (searchPJ == None):
                    for attrib in entity.GetAttributes():
                        Components[j][attrib.TagString] = attrib.TextString
                    j+=1


    for idx,port in enumerate(Module):

        if port != {}:
            if (idx + offset) < 9:
                if cardType == 'I':
                    wirestr = 'X4TERM0' + str(idx + 1 + offset)
                elif cardType == 'O':
                    wirestr = 'X1TERM0' + str(idx + 1 + offset)
            else:
                if cardType == 'I':
                    wirestr = 'X4TERM' + str(idx + 1 + offset)
                elif cardType == 'O':
                    wirestr = 'X1TERM' + str(idx + 1 + offset)
            if wirestr in port:
                wires.append(port[wirestr])
    print(wires)
        
    for comp in Components:
        for idx,wire in enumerate(wires):
            if ('X1TERM02' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['X1TERM02'] == wire):
                    if 'TAG1' in comp:
                        devices[idx] = comp['TAG1']
                    elif 'TAG1F' in comp:
                        devices[idx] = comp['TAG1F']
                    elif 'SIGCODE' in comp:
                        devices[idx] = comp['SIGCODE']
                    else:
                        devices[idx] = 'tag not found'
            if ('WIRENO' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['WIRENO'] == wire):
                    if 'TAG1' in comp:
                        devices[idx] = comp['TAG1']
                    elif 'TAG1F' in comp:
                        devices[idx] = comp['TAG1F']
                    elif 'SIGCODE' in comp:
                        devices[idx] = comp['SIGCODE']
                    else:
                        devices[idx] = 'tag not found'

            if ('X4TERM01' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['X4TERM01'] == wire):
                    if 'TAG1' in comp:
                        devices[idx] = comp['TAG1']
                    elif 'TAG1F' in comp:
                        devices[idx] = comp['TAG1F']
                    elif 'SIGCODE' in comp:
                        devices[idx] = comp['SIGCODE']
                    else:
                        devices[idx] = 'tag not found'

    print(devices)

    for idx in range(0, 31):
        if (idx + offset) < 9:
            descstr = 'DESCA0' + str(idx + 1 + offset)
        else:
            descstr = 'DESCA' + str(idx + 1 + offset)
        if descstr in Module[idx]:
            descriptions[idx] = Module[idx][descstr]
        else:
            descriptions[idx] = 'description not found'

    print(descriptions)
    
    for i in range(0,32):
        if (i < 10) & padzero.get():
            alias[i] = aliasPrefix +  str(slot) + cardType + indexPrefix + '0' + str(i) + indexPostfix
        else:
            alias[i] = aliasPrefix +  str(slot) + cardType + indexPrefix + str(i) + indexPostfix
        i+=1



    writer.writerow(['.3'])

    writer.writerow(['TYPE', 'SCOPE', 'NAME', 'DESCRIPTION', 'DATATYPE', 'SPECIFIER', 'ATTRIBUTES'])

    for i in range(0,31):
        if devices[i] != 0:
            print(devices[i])
            print(descriptions[i])
            print(alias[i])
            writer.writerow(['ALIAS', '', devices[i], descriptions[i], 'BOOL', alias[i], '(ExternalAccess := Read/Write)'])

    ofile.close()






top = Tk()
top.title('ACADE -> Studio5000 Tag Data Extractor')
top.geometry('1000x400')

padzero = BooleanVar()

C_pickModule = Combobox(top)
C_pickModule['values']= ('1769-IQ16', '1769-OB16', '5069-IB16', '5069-OB16', 'J-Block Inputs','J-Block Outputs')
C_pickModule.current(0)
C_pickModule.grid(column=1, row=0)

B_run = Button(top, text='Run', command=run)
B_run.grid(column=0, row=2)

L_chooseModule = Label(top, text='Choose Module:')
L_chooseModule.grid(column=0, row=0)

L_enterAliasPrefix = Label(top, text='Enter Alias Prefix:')
L_enterAliasPrefix.grid(column=0, row=1)

E_aliasPrefix = Entry(top, width=20)
E_aliasPrefix.grid(column=1, row=1)

L_enterSlot = Label(top, text='Slot')
L_enterSlot.grid(column=2, row=0)

C_slot = Combobox(top)
C_slot['values']= ('','0:','1:','2:','3:','4:','5:','6:','7:','8:','9:')
C_slot.current(0)
C_slot.grid(column=2, row=1)

L_cardType = Label(top, text='Card Type:')
L_cardType.grid(column=3, row=0)

C_type = Combobox(top)
C_type['values']= ('I','O')
C_type.current(0)
C_type.grid(column=3, row=1)

L_enterIndexPrefix = Label(top, text='Enter Index Prefix:')
L_enterIndexPrefix.grid(column=4, row=0)

E_indexPrefix = Entry(top, width=10)
E_indexPrefix.grid(column=4, row=1)

L_enterIndexPostfix = Label(top, text='Enter Index Postfix:')
L_enterIndexPostfix.grid(column=5, row=0)

E_indexPostfix = Entry(top, width=10)
E_indexPostfix.grid(column=5, row=1)

K_padZero = Checkbutton(top, text='Pad Zeroes', variable=padzero)
K_padZero.grid(column=6, row=1)

top.mainloop()

