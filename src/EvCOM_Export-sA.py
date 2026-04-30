# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_Export-sA.py
# Export Sa data WITHOUT THE EKVLOG or if the EKVLOG
#   is corrupted. This program resets the grid to GPS Distance (nmi).
#   If the EKVlog does not exist, use EV-COM_Export-Sa-Cell-Region_No-EKVlog.pl
#   This is because the output files will not be concatenated and the interval
#   values will not be unique.  Use the program Concatenate_EV-OutputFiles.pl to 
#   merge the output files.
# This program uses Echoview COM Scripting to read and export the
#   Sa values.
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
    # the output directory name
    odname = 'EK80_SA_WC'
    #*********************************************************
    
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)

    # get the list of export variables that we want to export
    #exportvars = evApp.getExportVars()

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
            gonogo1 = evApp.getEvExporters()
            if not gonogo1:
                evApp.closeEvFile(fl)
            else:
                evApp.EvFile.PreReadDataFiles
                exporters_count = evApp.getEvExportersCount()
                print(f'Number of Exporters: {exporters_count}')
                for i in range(exporters_count):
                    exporter_name = evApp.getEvExporterNamebyItem(i)
                    celloutfile = Path(str(fl.stem)+'_'+exporter_name)
                    celloutfile = outdir / celloutfile.with_suffix('.csv')
                    print(f'Exporting: {exporter_name} to {celloutfile}')
                    exported = evApp.exportEvExporterbyItem(i, celloutfile)
                    if (exported):
                        print('Successful Export')
                    else:
                        print('Unsuccessful Export')

        # close the EV file
        evApp.closeEvFile(fl)
    
    
    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
