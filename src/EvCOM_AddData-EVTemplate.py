# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# use COM & Echoview to create an EV file from a template
#  This program opens an existing EV file, gets the data file names, inputs those file
#  to the template, and imports lines
#  This includes the path to the data files
#
# jech
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

import win32com.client
from pathlib import Path
import EvCOM
import re

proname = 'EvCOM_AddData-EVTemplate.py'
numerrors = 0
# instantiate Echoview COM
evApp = EvCOM.Utilities(numerrors)

tmpfiles = evApp.getEvFiles('Select EV Files')
EVflist = list()
for f in tmpfiles[0]:
    EVflist.append(Path(f))

# put the filenames in a dictionary using the base part of the file name as the key
# the bottom line and region files should match these if they exist
fdict = {}
for f in EVflist:
    fdict[str(f.stem)] = {'EVfile': f}

# create an output directory for the new EV files
outdir = EVflist[0].parent / 'new_EV_Files'
evApp.createDir(outdir)

# create an error file for the error messages
evApp.Errors.createErrorFile(proname, outdir)

# get the seabed echo line files
tmpfiles = evApp.getevlFiles('Select EV Line Files')
for f in tmpfiles[0]:
    bname = re.search(r'd\d{8}_t\d{6}-t\d{6}', f)
    if (bname.group(0)):
        fdict[bname.group(0)]['evl'] = Path(f)
    else:
        fdict[bname.group(0)]['evl'] = False
# the name of the line that you will import. This needs to be a line that is in the 
# template
#EVlinename = 'seabed_echo'
EVlinename = 'ev bottom'

# get the region definition files
tmpfiles = evApp.getevrFiles('Select EV Region Definition Files')
for f in tmpfiles[0]:
    bname = re.search(r'd\d{8}_t\d{6}-t\d{6}', f)
    if (bname.group(0)):
        fdict[bname.group(0)]['evr'] = Path(f)
    else:
        fdict[bname.group(0)]['evr'] = False

# select the EV template
EVtemplate = Path(evApp.getEvFiles('Select EV Template')[0][0])

### start EV COM interface
#EvApp = win32com.client.Dispatch('EchoviewCom.EvApplication')

### 
for f in fdict.keys():
    # open the EV file
    evfl = fdict[f]['EVfile']
    print(f'Doing EV File: {evfl}')
    gonogo0 = evApp.openEvFile(evfl)
    if not gonogo0:
        evApp.closeEvFile(evfl)
    else:
        # get the filesets and data files for each fileset
        # number of file sets
        filesetcount = evApp.getFilesetCount()

    for i in range(filesetcount):
        filesetname = evApp.getFilesetNamebyIndex(i)
        datafilecount = evApp.getFilesetDataFileCountbyIndex(i)
        filesetfiles = evApp.getFilesetDataFilesbyIndex(i, datafilecount)
        if datafilecount >= 0:
            fdict[f].setdefault('fileset', {}).update({filesetname: 
                                                       {'idx': i,
                                                        'nf': datafilecount,
                                                        'fs': filesetfiles}
                                                       })

    # close the original EV file
    evApp.closeEvFile(evfl)

    # create the new EV file with the template
    newevfl = outdir / Path(str(evfl.stem)+'_new.EV')
    gonogo1 = evApp.createEvFile(EVtemplate)
    if not gonogo1:
        print('Unable to create new EV file')
    else:
        # add data files
        for fs in fdict[f]['fileset'].keys():
            # the number of data files
            fsdx = fdict[f]['fileset'][fs]['idx']
            nf = fdict[f]['fileset'][fs]['nf']
            if nf >= 1:
                dflist = fdict[f]['fileset'][fs]['fs']
                # convert to Path
                dflist = [Path(f) for f in dflist]
                dfpath = dflist[0].parent
                # clear the data path
                evApp.clearDataPaths()
                # add the data path
                evApp.addDataPaths(dfpath)
                # add the data files
                for df in dflist:
                    evApp.addDataFiles(fsdx, df)

        # add seabed echo lines
        if 'evl' in fdict[f]:
            linefile = fdict[f]['evl']
            evApp.importEvLine(EVlinename, linefile)

        # add regions
        if 'evr' in fdict[f]:
            regionfile = fdict[f]['evr']
            evApp.importRegionDefinitionsAll(regionfile)

        # save the EV file
        evApp.saveAsEvFile(newevfl)

        # close the EV file
        evApp.closeEvFile(newevfl)

# close the error file
evApp.Errors.closeErrorFile()
# close the Echoview COM
evApp.closeEvCom()

# end main

