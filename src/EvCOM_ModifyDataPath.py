# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_ModifyDataPath.py
#
# modify the data paths for files in an EV file
# useful for when data file locations change
#
# jech
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

import win32com.client
from pathlib import Path
import EvCOM

proname = 'EvCOM_ModifyDataPaths.py'
numerrors = 0
# instantiate Echoview COM
evApp = EvCOM.Utilities(numerrors)

# select the EV files
tmpfiles = evApp.getEvFiles('Select EV Files')
EVflist = []
for f in tmpfiles[0]:
    EVflist.append(Path(f))
EVdir = EVflist[0].parent

'''
if use glob
#for f in TopDirectory.glob('**/*_GPS.EV'):
for f in TopDirectory.glob('*.EV'):
    #print('EV file: ', f)
    EVflist.append(f)
'''

# modify the data directory for the files in each file set
newFilesetDataDirectory = []
for evfl in EVflist:
    print(f'Doing EV File: {evfl}')
    gonogo0 = evApp.openEvFile(evfl)
    if not gonogo0:
        evApp.closeEvFile(evfl)
    else:
        # clear all the data paths
        evApp.clearDataPaths()
        # get the filesets and data files for each fileset
        # number of filesets
        filesetcount = evApp.getFilesetCount()
        if (len(newFilesetDataDirectory) == 0):
            for i in range(filesetcount):
                # select the new data directory for each file set
                filesetname = evApp.getFilesetNamebyIndex(i)
                wtitle = 'Select the Data Directory for '+str(filesetname)
                newFilesetDataDirectory.append(evApp.selectDirectory(wtitle, EVdir))
                #print(f'dataDirectory: {newFilesetDataDirectory[i]}')
        for i in range(filesetcount):
            # add a data path
            evApp.addDataPaths(newFilesetDataDirectory[i])
        evApp.saveEvFile(evfl)
        evApp.closeEvFile(evfl)

    '''
    # do an open, save, close to generate the new .evi files if necessary
    evApp.openEvFile(evfl)
    evApp.saveEvFile(evfl)
    evApp.closeEvFile(evfl)
    '''

# close the Echoview COM
EvApp = None

# end main

