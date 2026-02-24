######################################################
# EvCom_Open-Save-Close.py
# Open, save, and close EV files. This is used when EV files were created
# with a lower version of Echoview and a new version has been upgraded
# this uses Echoview COM
#
# Original code by Victoria Price
# modifed by jech

#from PyQt5 import QtWidgets
import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_Open-Save-Close'
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)
    # get the files, this returns a tuple with the first list as a list of lists
    # with the filenames and the second list the file filter
    tmpfiles = evApp.getEVFiles()
    # tmpfiles is a tuple, convert to a list and format as pathlib Path
    ev_files = list()
    for f in tmpfiles[0]:
        ev_files.append(Path(f))
    
    # the output directory. Use the same directory as the original EV files
    outdir = ev_files[0].parent
    evApp.createDir(outdir)

    # create an error file for the error messages
    evApp.Errors.createErrorFile(proname, outdir)
    
    for f in ev_files:
        # open an EV file
        evApp.openEVFile(f)
        # save the EV file
        evApp.saveEVFile(f)
        # close the EV file
        evApp.closeEVFile(f)
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEVCom()


### end of main
