import win32com.client
import re
import csv
from tkinter import *
from tkinter.ttk import *
import codecs


#Defines different I/O modules to search for
#The key is the human-readable model number
#the value is the block name in AutoCAD
ModuleDefinitions = {
    '1769-IQ16': 'PLCIO_5E769',
    '1769-OB16': 'PLCIO_5EB55',
    '5069-IB16': 'PLCIO_5B719',
    '5069-OB16': 'PLCIO_5B2C5',
    'J-Block Inputs': 'PLCIO_59454',
    'J-Block Outputs': 'PLCIO_5980A',
    'EX600-DXPD (Amazon)': 'PLCIO_5F24F',
    'EX600_DYPB (Amazon)': 'PLCIO_6054F',
    'EX600-DXPD (Covance)' : 'HDV1_008'
    }

firstRun = 1;
def run():
    global firstRun;

    #get user input values from GUI
    aliasPrefix = E_aliasPrefix.get()
    slot = C_slot.get()
    cardType = C_type.get()
    indexPrefix = E_indexPrefix.get()
    indexPostfix = E_indexPostfix.get()
    
    #opens or creates output files for tags and parameters
    #user checkbox selects whether to append new tags to existing tagout or overwrite
    if (appendTags.get()):
        ofile  = open('tagout.csv', "a", newline='')
    else:
        ofile  = open('tagout.csv', "w", newline='')

    parfilestr = aliasPrefix
    parfilestr = re.sub('[^A-Za-z0-9]+', '', parfilestr)
    parfilestr = parfilestr + '.par'
    
    parfile = open(parfilestr, "w", newline='\n')
    writer = csv.writer(ofile)

    #finds AutoCAD in windows if running
    acad = win32com.client.Dispatch("AutoCAD.Application")

    #retreives the active drawing
    doc = acad.ActiveDocument

    #initialize empty array of dictionaries for up to 32 module ports
    Module = [{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},
              {},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}]
    #initialize empty array of dictionaries for up to 32 connected components
    Components = [{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},
                  {},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}]

    #initialize empty lists to save connected wire numbers, devices, etc.
    wires = []
    devices = [0] * 32
    descriptions = [''] * 32
    alias = [''] * 32
    ports = 16
    j=0


    #iterate over all objects in the document
    for entity in acad.ActiveDocument.ModelSpace:

        #save the name of the entity
        name = entity.EntityName

        #if the entitiy is a "Block"
        if name == 'AcDbBlockReference':

            HasAttributes = entity.HasAttributes

            EffectiveName = entity.EffectiveName

            #check if the block effective name matches the one the user selected
            if (HasAttributes & (EffectiveName == ModuleDefinitions[C_pickModule.get()])):
                
                foundFirst = False  #clear to false for next first port search
                offset = 0          #clear to zero for next first port search

                #for each attribute in the block
                for attrib in entity.GetAttributes():
                    i=0
                    found = False
                    while not found:
                        if i < 9:
                            searchstr = '\w+0' + str(i+ 1)      #search for anything ending with a 01, 02, 03...
                        else:
                            searchstr = '\w+' + str(i + 1)      #search for anything ending with a 10, 11, 12...

                        search = re.search(searchstr, attrib.TagString) #use string found above to see if the current tag matches (DESC01, DESC02...)

                        #if the search is successful
                        if search:
                            found = True
                            
                            #if this is the first port to be found, set offset
                            if not foundFirst:
                                offset = i          #set the offset to the first port
                                foundFirst = True   #prevent further searching

                            #set the 'i'th module array value to a tagstring key and textstring value pair (DESCA01:I.0/01)...
                            Module[i-offset][attrib.TagString] = attrib.TextString
                        else:
                            i+=1
                        if i > 64:
                            found = True    #break while loop if nothing found
                        
            #for every object besides the one selected by the user
            else:
                #create a bunch of search strings to get rid of common little blocks I am not interested in
                searchTerminal = re.search('HT0W01', entity.EffectiveName)
                searchPJ = re.search('HC01PJ_1', entity.EffectiveName)
                searchDest = re.search('HA1D', entity.EffectiveName)
                searchWD = re.match('WD', entity.EffectiveName)
                
                #if it is none of those things, save it to a components list
                if (searchTerminal == None) & (searchDest == None) & (searchWD == None) & (searchPJ == None):
                    for attrib in entity.GetAttributes():
                        Components[j][attrib.TagString] = attrib.TextString
                    j+=1

    #for each port on the module
    for idx,port in enumerate(Module):

        #if the port was found in the drawing and had values assigned
        if port != {}:
            #set search strings for the wire numbers attached to the port
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
            #if the wirestr tag exists then assign the corresponding wire number to the wires array
            if wirestr in port:
                wires.append(port[wirestr])
    print(wires)

    #for all 32 ports, save description values
    for idx in range(0, 31):
        if (idx + offset) < 9:
            descstr = 'DESCA0' + str(idx + 1 + offset)
        else:
            descstr = 'DESCA' + str(idx + 1 + offset)
        if descstr in Module[idx]:
            descriptions[idx] = Module[idx][descstr]        #save the description to the descriptions list
        else:
            descriptions[idx] = 'description not found'

    print(descriptions)

    #for each of the found components that were not the module and were not filtered out
    for comp in Components:
        for idx,wire in enumerate(wires):
            #if the wire number from the module port is also on a terminal of the component and is also not 200 or 202
            #save the tag value to the devices list
            if ('X1TERM02' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['X1TERM02'] == wire):
                    devices[idx] = findComponentName(comp)
                    if (devices[idx] != 'tag not found'):
                        if ((descriptions[idx] == 'description not found') | (descriptions[idx] == '')):
                            if ('DESC1' in comp):
                                descriptions[idx] = comp['DESC1']
            if ('WIRENO' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['WIRENO'] == wire):
                    devices[idx] = findComponentName(comp)
                    if ((descriptions[idx] == 'description not found') | (descriptions[idx] == '')):
                        if ('DESC1' in comp):
                            descriptions[idx] = comp['DESC1']

            if ('X4TERM01' in comp):
                if (wire != '202') & (wire !='200') & (wire != '') & (comp['X4TERM01'] == wire):
                    devices[idx] = findComponentName(comp)
                    if ((descriptions[idx] == 'description not found') | (descriptions[idx] == '')):
                        if ('DESC1' in comp):
                            descriptions[idx] = comp['DESC1']
                
    print(devices)

    #set the alias names according to the user input values
    for i in range(0,32):
        if (i < 10) & padzero.get():
            alias[i] = aliasPrefix +  str(slot) + cardType + indexPrefix + '0' + str(i) + indexPostfix
        else:
            alias[i] = aliasPrefix +  str(slot) + cardType + indexPrefix + str(i) + indexPostfix
        i+=1

    #if this is the first time the run button has been pressed, or if the append option is not selected
    #write header values for the studio5000 tag import format

    if (firstRun or not appendTags.get()):
        writer.writerow(['.3'])
        writer.writerow(['TYPE', 'SCOPE', 'NAME', 'DESCRIPTION', 'DATATYPE', 'SPECIFIER', 'ATTRIBUTES'])

    firstRun = 0    #this is no longer the first run
    #write the first two lines of the parameter file
    k = 1
    parfile.write("#" + str(k) + "=" + C_pickModule.get() + "\n")
    k += 1
    parfile.write("#" + str(k) + "=" + "Status" + "\n")
    k += 1

    #final write-to-file stage
    for i in range(0,31):
        print(devices[i])
        print(descriptions[i])
        print(alias[i])

        if devices[i] != 0:
            #write all the information to the csv file for each tag
            writer.writerow(['ALIAS', '', devices[i], descriptions[i], 'BOOL', alias[i], '(ExternalAccess := Read/Write)'])
        #write the PLC tag for each parameter
        parfile.write("#" + str(k) + "={::[PLC]" + alias[i] + "}\n")
        k += 1

        #write the description for each parameter
        if (descriptions[i] == ''):
            parfile.write("#" + str(k) + "=" + "SPARE\n")
        else:
            parfile.write("#" + str(k) + "=" + descriptions[i] + "  (" + str(devices[i]) + ")" + "\n")
        k += 1
            
    #close the files
    ofile.close()
    parfile.close()

    with open(parfilestr,'r') as file:
        filedata = file.read()

    filedata = filedata.replace(' ',"\u00A0")           #replace all spaces with no-break space unicode value
    filedata = filedata.replace('\n',"\u000d\u000a")    #replace all newline characters with carriage return, line feed

    #write as binary in utf-16 little endian (UCS-2 LE) with an explicit BOM
    with open(parfilestr, 'wb') as file:
        file.write(codecs.BOM_UTF16_LE)
        file.write(filedata.encode('utf-16-le'))


def findComponentName(comp):
        if 'SIGCODE' in comp:
            if comp['SIGCODE'] != '':
                deviceName = comp['SIGCODE']                    
        if 'TAG2' in comp:
            if comp['TAG2'] != '':
                deviceName = comp['TAG2']
        if 'TAG1F' in comp:
            if comp['TAG1F'] != '':
                deviceName = comp['TAG1F']
        if 'TAG1' in comp:
            if comp['TAG1'] != '':
                deviceName = comp['TAG1']
        if deviceName == '':
            deviceName = 'tag not found'
        return deviceName



#Tkinter GUI code
top = Tk()
top.title('ACADE -> Studio5000 Tag Data Extractor')
top.geometry('1000x400')

padzero = BooleanVar()
appendTags = BooleanVar()

C_pickModule = Combobox(top)
C_pickModule['values'] = list(ModuleDefinitions.keys())
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

K_appendTags = Checkbutton(top, text='Append taglist', variable=appendTags)
K_appendTags.grid(column=6, row=2)

top.mainloop()

