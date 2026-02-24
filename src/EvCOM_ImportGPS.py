# -*- coding: utf-8 -*-
#+++++++++++++++++++++++++++++++++++++++++++++++
# EvCom_ImportGPS.py
# import GPS .csv files into EV files
# uses Echoview COM
#
# Original code by Victoria Price
# modified by jech
#+++++++++++++++++++++++++++++++++++++++++++++++

import EvCOM
from pathlib import Path
import re


if __name__=='__main__':
    ###
    # variables to modify
    numerrors = 0
    proname = 'EvCOM_ImportGPS'
    
    # instantiate Echoview COM
    evApp = EvCOM.Utilities(numerrors)
    # get the files, this returns a tuple with the first list as a list of lists
    # with the filenames and the second list the file filter
    tmpfiles = evApp.getEvFiles()
    # tmpfiles is a tuple, convert to a list and format as pathlib Path
    ev_files = list()
    for f in tmpfiles[0]:
        ev_files.append(Path(f))
    
    # get the gps files
    tmpfiles = evApp.getCSVFiles()
    gps_files = list()
    for f in tmpfiles[0]:
        gps_files.append(Path(f))
    
    # create an error file for the error messages in the same directory as the
    # EV files
    evApp.Errors.createErrorFile(proname, ev_files[0].parent)
    
    for fl in ev_files:
        foundmatch = False
        # open an EV file
        gonogo0 = evApp.openEvFile(fl)
        if not gonogo0:
            evApp.closeEvFile(fl)
        else:
            # add the GPS fileset
            gonogo1 = evApp.addFileset('GPS')
            if not gonogo1:
                evApp.closeEvFile(fl)
            else:
                # add data files to the fileset
                # the GPS should have a date in the file name. Match this to
                # the date of the EV file
                evdate = fl.name.split('_')[0]
                evdate = evdate[1:]
                evetime = fl.name.split('-')[1]
                evetime = evetime[1:7]
                idx = 0
                for gfl in gps_files:
                    gpsdate = gfl.name.split('_')[1]
                    gpsdate = gpsdate.split('-')[0]
                    if re.match(evdate, gpsdate):
                        foundmatch = True
                        if int(evetime) < 120000:
                            gflist = gfl
                        else:
                            gflist = [gfl, gps_files[idx+1]]
                        gonogo2 = evApp.addDataFiles('GPS', gflist)
                        if not gonogo2:
                            #print('was not able to add files, closing')
                            evApp.closeEvFile(fl)
                            break
                        else:
                            #print('found a match, saving and closing')
                            evApp.saveEvFile(fl)
                            evApp.closeEvFile(fl)
                            break
                    idx += 1
                if not foundmatch: 
                    #print('no match, saving and closing')
                    evApp.saveEvFile(fl)
                    evApp.closeEvFile(fl)        

    # close the error file
    evApp.Errors.closeErrorFile()
    # close the Echoview COM
    evApp.closeEvCom()

### end of main
