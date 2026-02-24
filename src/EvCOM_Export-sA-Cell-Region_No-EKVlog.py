# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_Export-sA-Cell-Region_No-EKVlog.py
# Export Sa data by cells and regions for data files WITHOUT THE EKVLOG or if the EKVLOG
#   is corrupted. This program resets the grid to GPS Distance (nmi).
#   If the EKVlog does not exist, use EV-COM_Export-Sa-Cell-Region_No-EKVlog.pl
#   This is because the output files will not be concatenated and the interval
#   values will not be unique.  Use the program Concatenate_EV-OutputFiles.pl to 
#   merge the output files.
# This program uses Echoview COM Scripting to read and export the
#   Sa values.
# Because importing Sa data to Oracle is currently a bit complex,
#   I export the Sa data to csv files and then import to Oracle
#   in a separate step.
# The exported variable is currently hardcoded, and the file labels
#   are based on this
# Echoview scripting creates pre-defined output file names, but
#   modify them to our standard filename protocol.
# Actually, this program is really hardcoded for variables, output,
#   etc...
# Program Flow:
# -Get the EV files to export
# -Echoview scripting module is activated and exports .csv files
#
# Original code by Victoria Price
# translated from original perl programs
# modified by jech
#+++++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path


if __name__=='__main__':
    numerrors = 0
    proname = 'EvCOM_Export-sA'
    #*********************************************************
    # variables to modify
    # the Sv echogram to get the line from
    svechogram = 'Sv raw pings T5'
    # export_mode will set whether to output multiple files or a single file per EV
    # This is set in the EVfile -> Export dialog box; 
    # 1 = database (multiple files), 2 = spreadsheet (single file)
    export_mode = '1'
    # This sets the time/distance grid to GPS (nmi)
    ETimeDistanceGridMode = '2'
    # This sets the elementary distance sampling unit
    EDSU = 0.5
    # output empty cells. This pads the data with zeros
    emptycell = True
    # the output directory name
    odname = 'EV_Sa_Files'
    #*********************************************************
    
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)

    # get the list of export variables that we want to export
    exportvars = evApp.getExportVars()

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
        celloutfile = Path(str(fl.stem)+'-cells')
        celloutfile = outdir / celloutfile.with_suffix('.csv')
        regionoutfile = Path(str(fl.stem)+'-regions')
        regionoutfile = outdir / regionoutfile.with_suffix('.csv')
        # open an EV file
        gonogo0 = evApp.openEvFile(fl)
        if not gonogo0:
            evApp.closeEvFile(fl)
        else:
            # make sure we export the correct variables
            # to do this we first disable all export variables
            evApp.enableExportVariables('D', 'ALL')
            # need to enable the mandatory Echoview export variables that were
            # disabled and these can only be done using the command interface
            # command
            evApp.enableMandatoryExportVariables()
            # then we enable only those that we need
            evApp.enableExportVariables('E', exportvars)
            # set empty cell output
            evApp.EvFile.Properties.Export.EmptyCells = emptycell
            # set database or single file output
            evApp.EvFile.Properties.Export.Mode = export_mode
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
                    # set the time/distance grid to GPS distance (nmi)
                    evApp.setTimeDistanceGrid(ETimeDistanceGridMode, EDSU)
                    # save the changes
                    evApp.saveEvFile(fl)
                    # preread the data files
                    evApp.EvFile.PreReadDataFiles
                    # export the Sv/sa data by cells
                    evApp.exportIntegrationByCells(celloutfile)
                    # export the Sv/sa data by regions by cells
                    evApp.exportIntegrationByRegionsByCells(regionoutfile)
        # close the EV file
        evApp.closeEvFile(fl)
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
