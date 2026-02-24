# -*- coding: utf-8 -*-
#++++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_ExportVesselSpeed.py
# export vessel speed from an EV file
#
# Original code by Victoria Price
#
# modified by jech
#++++++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_ExportVesselSpeed'
    #**********************************************
    # variables to modify
    # the Sv echogram to get the line from
    svechogram = 'VS'
    # the file name extension for the output files
    fnext = '_VS.csv'
    # the file name prefix for the output files
    fnpre = 'HB202205_'
    # the name of the output directory
    odname = 'EV_vesselspeed'
    #**********************************************
    
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
                    outfile = Path(fnpre+str(fl.stem)+fnext)
                    outfile = outdir / outfile
                    # export the line
                    gonogo4 = evApp.exportEvData(outfile)
        # close the EV file
        evApp.closeEvFile(fl)
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
