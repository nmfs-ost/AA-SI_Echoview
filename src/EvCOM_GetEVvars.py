# -*- coding: utf-8 -*-
#++++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_GetEVvars.py
#
# get the variable names from an EV file
#
# jech
#++++++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_GetEVvars'
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)
    # get the files, this returns a tuple with the first list as a list of lists
    # with the filenames and the second list the file filter
    tmpfiles = evApp.getEvFiles()
    # tmpfiles is a tuple, convert to a list and format as pathlib Path
    ev_files = list()
    for f in tmpfiles[0]:
        ev_files.append(Path(f))
    
    # create the output directory
    outdir = ev_files[0].parent
    evApp.createDir(outdir)

    # create an error file for the error messages
    evApp.Errors.createErrorFile(proname, outdir)
    
    # open an EV file
    evApp.openEvFile(ev_files[0])

    # get variable names
    varnames = evApp.getVarNames(ev_files[0])

    # close the EV file
    evApp.closeEvFile(ev_files[0])
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
