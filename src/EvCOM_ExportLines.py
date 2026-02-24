# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_ExportLines.py
# export lines from an EV file
#
# Original code by Victoria Price
# modified by jech
#+++++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_ExportLines'
    #*****************************************************
    # variables to modify
    # the Sv echogram to get the line from
    #svechogram = 'Sv wideband pings T1'
    svechogram = 'Sv pings T4'
    # the line name
    #linename = 'bottom line'
    #linename = 'seabed_echo'
    linename = 'ev bottom'
    # the file name extension for the output files
    fnext = '_evseabed'
    odname = 'EV_seabedlines'
    #*****************************************************
    
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)

    # get the files, this returns a tuple with the first list as a list of lists
    # with the filenames and the second list the file filter
    tmpfiles = evApp.getEvFiles('Select EV Files')
    # tmpfiles is a tuple, convert to a list and format as pathlib Path
    ev_files = list()
    for f in tmpfiles[0]:
        ev_files.append(Path(f))
    
    # create the output directory
    outdir = ev_files[0].parent.parent / odname
    evApp.createDir(outdir)

    # create an error file for the error messages
    evApp.Errors.createErrorFile(proname, outdir)
    
    for fl in ev_files:
        # open an EV file
        gonogo0 = evApp.openEvFile(fl)
        if not gonogo0:
            evApp.closeEvFile(fl)
        else:
            # select the echogram
            gonogo1 = evApp.getEvEchogram(svechogram)
            if not gonogo1:
                evApp.closeEvFile(fl)
            else:
                # make sure it is an acoustic variable
                gonogo2 = evApp.asVarAcoustic(svechogram)
                if not gonogo2:
                    evApp.closeEvFile(fl)
                else:
                    # get the line
                    gonogo3 = evApp.getEvLine(linename)
                    if not gonogo3:
                        evApp.closeEvFile(fl)
                    else:
                        outfile = Path(str(fl.stem)+fnext)
                        outfile = outdir / outfile
                        # export the line
                        gonogo4 = evApp.exportEvLine(linename, outfile)
        # close the EV file
        evApp.closeEvFile(fl)
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
