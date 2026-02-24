# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_ExportGPS.py
# export GPS GPGGA from an EV file
#
# Original code by Victoria Price
# modified by jech
#+++++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_ExportGPS'
    #****************************************************
    # variables to modify
    # the Sv echogram to get the line from
    svechogram = 'position GPS fixes GPGGA'
    # the file name extension for the output files
    fnext = '_GPS.csv'
    # file name prefix
    fnpre = 'DN202203_'
    # output directory name
    odname = 'EV_GPS'
    #****************************************************
    
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
    outdir = ev_files[0].parent / odname
    evApp.createDir(outdir)

    # create an error file for the error messages
    evApp.Errors.createErrorFile(proname, outdir)
    
    for fl in ev_files:
        # open an EV file
        print('Doing: ', str(fl))
        gonogo0 = evApp.openEvFile(fl)
        if not gonogo0:
            evApp.closeEvFile(fl)
        else:
            # select the echogram
            gonogo1 = evApp.getEvVarName(svechogram)
            if not gonogo1:
                evApp.closeEvFile(fl)
            else:
                outfile = Path(fnpre+str(fl.stem)+fnext)
                outfile = outdir / outfile
                # export the GPS data
                gonogo4 = evApp.exportEvGPS(outfile)
        # close the EV file
        evApp.closeEvFile(fl)
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
