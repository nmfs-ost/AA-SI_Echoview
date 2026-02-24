################################################################
# EvCOM.py
# main set of classes and modules to interact with Echoview via COM
#
# Original code by Victoria Price
# modified by jech

from PyQt5  import QtWidgets, QtCore
import win32com.client
from pathlib import Path
import pythoncom

class Utilities:
    def __init__(self, nerrors):
        '''Instantiate Echoview COM'''
        #print('numerrors: ', nerrors)
        #super(Utilities, self).__init__(self)
        self.app = QtWidgets.QApplication([])
        print ("Program Loading and Compiling")
        print ("Please Wait.....")
        self.Errors = Errors()
        #self.app=self.Base.app
        ### Open the Echoview Scripting Module
        self.numerrors = nerrors
        try:
           self.evo=win32com.client.Dispatch('EchoviewCom.EvApplication')
           print('Initiated Echoview COM')
        except:
            print("Failed to Initiate Echoview COM")
            self.numerrors += 1
            pass

    def selectDirectory(self, wcaption, startdir):
        ''' select an existing directory using a Qt dialog
            input: the window caption
        '''
        selected_directory = QtWidgets.QFileDialog.getExistingDirectory(parent=None,
                                        caption=wcaption, directory=str(startdir))
        if len(selected_directory) <= 0:
            self.evo.Quit()
        else:
            return selected_directory

    def getevlFiles(self, wcaption):
        '''get the evl filenames using a Qt dialog'''
        evl_filenames = QtWidgets.QFileDialog.getOpenFileNames(parent=None, 
                                      caption=wcaption, filter="*.evl")

        self.nfiles = len(evl_filenames)
        #print('Number of files: ', self.nfiles)
        if self.nfiles <= 0:
            self.evo.Quit()
        else:
            return evl_filenames

    def getevrFiles(self, wcaption):
        '''get the evr filenames using a Qt dialog'''
        evr_filenames = QtWidgets.QFileDialog.getOpenFileNames(parent=None, 
                                      caption=wcaption, filter="*.evr")

        self.nfiles = len(evr_filenames)
        #print('Number of files: ', self.nfiles)
        if self.nfiles <= 0:
            self.evo.Quit()
        else:
            return evr_filenames

    def getcsvFiles(self, wcaption):
        '''get the csv filenames using a Qt dialog'''
        csv_filenames = QtWidgets.QFileDialog.getOpenFileNames(parent=None, 
                                      caption=wcaption, filter="*.csv")

        self.nfiles = len(csv_filenames)
        #print('Number of files: ', self.nfiles)
        if self.nfiles <= 0:
            self.evo.Quit()
        else:
            return csv_filenames

    def getEvFiles(self, wcaption):
        '''get the EV filenames using a Qt dialog and remove file names with
        *backup* in them'''
        ev_filenames = QtWidgets.QFileDialog.getOpenFileNames(parent=None, 
                                      caption=wcaption, filter="*.EV")

        ### remove the *(backup).EV  files from the list
        for  filename in ev_filenames:
            if "backup" in filename:
                ev_filenames.remove(filename)
        
        self.nfiles = len(ev_filenames)
        #print('Number of files: ', self.nfiles)
        if self.nfiles <= 0:
            self.evo.Quit()
        else:
            return ev_filenames

    def openEvFile(self, file):
        '''Open an Echoview EV file and create the file-level object. This is
           stored internally (i.e., self.EvFile) and used by the other methods.
           Input:   EV file name'''
        if isinstance(file, Path):
            file = str(file) 
        try:
            self.EvFile = self.evo.OpenFile(file)
            return True
        except pythoncom.com_error as error:
            print ('Error Opening EV File', file)
            a = str(error).split(',')
            error_message='Error Opening File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
        

    def createEvFile(self, evtemplate):
        '''Create an Echoview EV file and create the file-level object. This is
           stored internally (i.e., self.EvFile) and used by the other methods.
           Input: EV template name'''
        if isinstance(evtemplate, Path):
            file = str(evtemplate) 
        try:
            self.EvFile = self.evo.NewFile(evtemplate)
            return True
        except pythoncom.com_error as error:
            print('Error Creating EV File with template: ', evtemplate)
            a = str(error).split(',')
            error_message='Error Creating EV File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def saveEvFile(self, file):
        '''Save an Echoview EV file
           Input:   File name to save
                    EV file name'''
        try:
            self.EvFile.Save()
            print('Saving EV File: ', file)
            return True
        except pythoncom.com_error as error:
            print('Error Saving EV File: ', file)
            a = str(error).split(',')
            error_message='Error Saving File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def saveAsEvFile(self, fname):
        '''Save an Echoview file with a new name'''
        if isinstance(fname, Path):
            fname = str(fname) 
        try:
            self.EvFile.SaveAs(fname)
            print('SaveAs EV File: ', fname)
            return True
        except pythoncom.com_error as error:
            print('Error SaveAs EV File: ', fname)
            a = str(error).split(',')
            error_message='Error SaveAs File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
            

    def closeEvFile(self, file):
        '''Close an Echoview EV file
           Input:   EV file name'''
        if isinstance(file, Path):
            file = str(file) 
        try:
            self.evo.CloseFile(self.EvFile)
            print('Closing EV File: ', file)
            return True
        except pythoncom.com_error as error:
            print('Error Closing EV File: ', file)
            a = str(error).split(',')
            error_message='Error Closing File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def closeEvCom(self):
        '''Close the Echoview COM'''
        self.evo.Quit()


    def createDir(self, dirtocreate):
        '''Create a subdirectory, usually for output
           Input:   directory name to create'''
        if not isinstance(dirtocreate, Path):
            dirtocreate = Path(dirtocreate)
        if (dirtocreate.exists()):
            print('  Directory %s exists' % dirtocreate)
        else:
            try:
                dirtocreate.mkdir()
            except OSError:
                print('  Unable to create directory %s' % dirtocreate)
                exit()
            else:
                print('  Created directory %s' % dirtocreate)
    

    def getVarNames(self, file):
        '''Get a list of the variable names in an EV file
           Input: EV file name'''
        varNames = []
        if isinstance(file, Path):
            file = str(file) 
        try:
            vars = self.EvFile.Variables
            print('Getting variables from EV File: ', file)
            count = vars.count
            for i in range(count):
                varNames.append(vars[i].Name)
            return varNames
        except pythoncom.com_error as error:
            print('Error getting variables from EV File: ', file)
            a = str(error).split(',')
            error_message='Error getting variables from File: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def getEvVarName(self, varname):
        '''Get the variable by name in an EV file
           Input: variable name
        '''
        try:
            self.EvVarName = self.EvFile.Variables.FindByName(varname)
            return True
        except pythoncom.com_error as error:
            print('Error getting variable ', varname)
            a = str(error).split(',')
            error_message='Error getting variable '+varname+a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
            

    def getEvEchogram(self, egname):
        '''select the echogram that will be used to process
           Input: echogram name'''
        try:
            self.EvEchogram = self.EvFile.Variables.FindByName(egname)
            print('Getting Echogram: ', egname)
            return True
        except pythoncom.com_error as error:
            print('Error getting Echogram: ', egname)
            a = str(error).split(',')
            error_message='Error getting Echogram: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def asVarAcoustic(self, egname):
        '''confirm that the echogram is an acoustic variable
           Input: echogram name'''
        try:
            self.EvVarAcoustic = self.EvEchogram.AsVariableAcoustic()
            print('Acoustic Variable: ', egname)
            return True
        except pythoncom.com_error as error:
            print('Error Acoustic Variable: ', egname)
            a = str(error).split(',')
            error_message='Error Acoustic Variable: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
        

    def getEvLine(self, lname):
        '''Get the Echoview line
           Input: line name'''
        try:
            self.EvLine = self.EvFile.Lines.FindByName(lname)
            print('Getting line: ', lname)
            return True
        except pythoncom.com_error as error:
            print('Error Getting Line: ', lname)
            a = str(error).split(',')
            error_message='Error Getting Line: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def exportEvLine(self, lname, outfn):
        '''Export an Echoiew Line to both .csv and .evl format files
           Input:  line name
                   output filename
           The output filename does not need the extension/suffix. The first
           export is to .csv and the second to .evl formats.''' 
        try:
            # export as .csv
            outfn = outfn.with_suffix('.csv')
            self.EvVarAcoustic.ExportLine(self.EvLine, str(outfn), -1, -1)
            # export as .evl
            outfn = outfn.with_suffix('.evl')
            self.EvVarAcoustic.ExportLine(self.EvLine, str(outfn), -1, -1)
            print('Exporting line: ', lname)
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Line: ', lname)
            a = str(error).split(',')
            error_message='Error Exporting Line: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def importEvLine(self, lname, lfile):
        '''import a line to the EV file
           This is a very complicated method'''
        if isinstance(lfile, Path):
            lfile = str(lfile)
        try:
            self.EvFile.Import(lfile)
            ## you can't specify the name of the line, so you can either just import and
            ## the imported line will be the last line in list, or you can overwrite the
            ## specified line and delete the imported line
            importlineindex = self.EvFile.Lines.Count-1
            EvImportLine = self.EvFile.Lines.Item(importlineindex)
            importlinename = self.EvFile.Lines.Item(importlineindex).Name
            evlinename = self.EvFile.Lines.FindByName(lname)
            evlinename.OverwriteWith(EvImportLine)
            self.EvFile.Lines.Delete(EvImportLine)
            return True
        except pythoncom.com_error as error:
            print('Error Importing Line: ', lname)
            a = str(error).split(',')
            error_message='Error Importing Line: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def countEvLines(self):
        '''get the number of lines in the EV file'''
        nlines = self.EvFile.Lines.Count
        return nlines


    def getEvLineItem(self, ldx):
        '''get the item number of the line'''
        itemdx = self.EvFile.Lines.Item(ldx)
        return itemdx


    def getEvLinebyIndex(self, ldx):
        '''get the line name of the selected line'''
        itemname = self.EvFile.Lines.Item(ldx).Name
        return itemname


    def getEvLinebyName(self, lname):
        '''get the line name of the selected line'''
        try:
            ldx = self.EvFile.Lines.FineByName(lname)
            return ldx
        except pythoncom.com_error as error:
            print('Error Finding Line: ', lname)
            a = str(error).split(',')
            error_message='Error Finding Line: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def overwriteEvLine(self, oldline, newline):
        '''overwrite a line'''
        self.evo.OverwriteWith(oldline)
        return True
        

    def deleteEvLine(self, ldx):
        '''delete a line'''
        self.EvFile.Lines.Delete(ldx)
        return True


    def exportRegionDefinitionsAll(self, outfn):
        '''Export Echoview region definitions
           Input:  output name
           The output filename does not need the extension/suffix.
        '''
        try:
            # export as .csv
            outfn = outfn.with_suffix('.evr')
            self.EvFile.Regions.ExportDefinitionsAll(str(outfn))
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Region Defintions')
            a = str(error).split(',')
            error_message='Error Exporting Region Definitions'
            Errors.AppendErrorFile(self, error_message)
            return False


    def importRegionDefinitionsAll(self, rfile):
        '''import regions'''
        if isinstance(rfile, Path):
            rfile = str(rfile)
        try:
            self.EvFile.Import(rfile)
            return True
        except pythoncom.com_error as error:
            print('Error Importing Regions: ', rfile)
            a = str(error).split(',')
            error_message='Error Importing Regions: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def exportEvData(self, outfn):
        '''Export data from an echogram
           Input: output file name'''
        try:
            self.EvVarAcoustic.ExportData(str(outfn))
            print('Exporting Data to : ', str(outfn))
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Data')
            a = str(error).split(',')
            error_message='Error Exporting Data: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
        

    def exportEvGPS(self, outfn):
        '''Export GPS from an echogram
           Input: output file name'''
        try:
            self.EvVarName.ExportData(str(outfn),-1,-1)
            print('Exporting Data to : ', str(outfn))
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Data')
            a = str(error).split(',')
            error_message='Error Exporting Data: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False
        

    def getExportVars(self):
        '''Get the list of export variables'''
        # this mimics the perl command qw
        varlist = '''Date_E Date_M Date_S
                   Dist_E Dist_M Dist_S
                   Lat_E Lat_M Lat_S
                   Lon_E Lon_M Lon_S
                   Num_intervals Num_layers
                   Ping_E Ping_M Ping_S
                   Time_E Time_M Time_S
                   VL_end VL_mid VL_start
                   Depth_mean Height_mean
                   Layer_depth_max Layer_depth_min
                   Good_samples EV_filename Processing_date
                   Processing_time Program_version NASC PRC_NASC
                   Sv_mean Minimum_integration_threshold
                   Minimum_Sv_threshold_applied'''.split()
        return varlist


    def getMandatoryExportVars(self):
        '''Get the list of Echoview mandatory export variables. This needs to
           be called if you disabled all the export variables using the 
           command interface command in enableExportVariables. These variables can only
           be enabled using the command interface command'''
        # this mimics the perl command qw
        mvarlist = '''RegionID RegionName RegionClass
                   ProcessID Interval Layer LayerRangeMin LayerRangeMax
                   FirstLayerRangeStart LastLayerRangeStop
                   CGoodSamples CBadDataNoDataSamples CNoDataSamples
                   CBadDataEmptyWaterSamples'''.split()
        return mvarlist


    def enableMandatoryExportVariables(self):
        '''Enable the mandatory export variables that were disabled'''
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| RegionID")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| RegionName")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| RegionClass")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| ProcessID")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| Interval")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| Layer")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| LayerRangeMin")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| LayerRangeMax")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| FirstLayerRangeStart")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| LastLayerRangeStop")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| CGoodSamples")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| CBadDataNoDataSamples")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| CNoDataSamples")
        self.evo.Exec("Ev File | ExportAnalysisVariables +=| CBadDataEmptyWaterSamples")
        return True        


    def enableExportVariables(self, eord, vlist):
        '''Enable or disable a list of export variables in an EV file
           Syntax: Enable_EV_ExportVariables(EvFile, ['D' or 'E'], @listofvariables)
            where Evfile is the file to import the line,
            ['D' or 'E'] is 'D' for disable and 'E' for enable,
            and listofvariables is a list of the variable names where
            listofvariables can be a list or "ALL".  If "ALL" is specified
            all export variables will be enabled/disabled.
           Examples: 
            Disable all variables: Enable_EV_ExportVariables(Evfile, 'D', "ALL")
            Enable a list: Enable_EV_ExportVariables(Evfile, 'E', varlist)'''
        if vlist=='ALL' and eord=='D':
            # disable all export variables
            print('Disable All Export Variables')
            try:
                self.evo.Exec("Ev File | ExportAnalysisVariables =| None")
                return True
            except pythoncom.com_error as error:
                print('Error Disabling All Export Variables: ', eord)
                a = str(error).split(',')
                error_message='Error Disabling All Export Variables: '+ a[4]
                Errors.AppendErrorFile(self, error_message)
                return False
        elif vlist=='ALL' and eord=='E':
            # enable all export variables
            print('Enable All Export Variables')
            try:
                self.evo.Exec("Ev File | ExportAnalysisVariables =| All")
                return True
            except pythoncom.com_error as error:
                print('Error Enabling All Export Variables: ', eord)
                a = str(error).split(',')
                error_message='Error Enabling All Export Variables: '+ a[4]
                Errors.AppendErrorFile(self, error_message)
                return False
        else:
            # cycle through the export variable list to enable/disable
            if eord == 'E':
                print('enable vlist')
                for v in vlist:
                    self.EvFile.Properties.Export.Variables.Item(v).Enabled = True
            if eord == 'D':
                print('disable vlist')
                for v in vlist:
                    self.EvFile.Properties.Export.Variables.Item(v).Enabled = False
        return True


    def setTimeDistanceGrid(self, mode, edsu):
        '''set the time/distance grid mode and edsu
           Input:  mode: 2 = GPS (nmi)
                   edsu is the elementary sampling distance unit'''
        try:
            self.EvVarAcoustic.Properties.Grid.setTimeDistanceGrid(mode, edsu)
            return True
        except pythoncom.com_error as error:
            print('Error Setting Time/Distance Grid: ', mode)
            a = str(error).split(',')
            error_message='Error Setting Time/Distance Grid: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def exportIntegrationByCells(self, outfl):
        '''Export the Sv/sa integration data by cells. Currently this exports
           all cells.
           Input: cell output file name'''
        if isinstance(outfl, Path):
            outfl = str(outfl)
        try:
            self.EvVarAcoustic.ExportIntegrationByCellsAll(outfl)
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Integration by Cell: ', outfl)
            a = str(error).split(',')
            error_message='Error Exporting Integration by Cell: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def exportIntegrationByRegionsByCells(self, outfl):
        '''Export the Sv/sa integration data by regions by cells. Currently this
           exports all regions.
           Input: region output file name'''
        if isinstance(outfl, Path):
            outfl = str(outfl)
        try:
            self.EvVarAcoustic.ExportIntegrationByRegionsByCellsAll(outfl)
            return True
        except pythoncom.com_error as error:
            print('Error Exporting Integration by Region by Cells: ', outfl)
            a = str(error).split(',')
            error_message='Error Exporting Integration by Regions by Cell: '+ a[4]
            Errors.AppendErrorFile(self, error_message)
            return False


    def getFilesetCount(self):
        '''Get the number of filesets in an EV file'''
        self.filesetcount = self.EvFile.Filesets.Count
        return self.filesetcount


    def getFilesetNamebyIndex(self, idx):
        '''Get the name of the fileset.
           Input: the index of the fileset
           Output: the name of the fileset'''
        filesetname = self.EvFile.Filesets.Item(idx).Name
        return filesetname


    def getFilesetDataFileCountbyIndex(self, idx):
        '''Get the number of data files in the fileset.
           Input: the index of the fileset
           Output: the number of data files in the fileset'''
        datafilecount = self.EvFile.Filesets.Item(idx).DataFiles.Count
        return datafilecount


    def getFilesetDataFilesbyIndex(self, idx, n):
        '''Get the file names of data files in the fileset.
           Input: the index of the fileset
                  the number of files to get
           Output: list of the data file names in the fileset'''
        filesetfiles = []
        for j in range(n):
            # get the filenames
            filesetfiles.append(self.EvFile.Filesets.Item(idx).DataFiles.Item(j).FileName)
        # convert to Path object
        filesetfiles = [Path(f) for f in filesetfiles]
        return filesetfiles


    def addFileset(self, fsname):
        '''Add a fileset to the EV file
           Input: Fileset name'''
        if isinstance(fsname, Path):
            fsname = str(fsname)
        self.EvFile.Filesets.Add(fsname)
        return True


#    def clearFilesetDataPathbyIndex(self, idx):
#        '''Clear the datapaths from a file set'''
#        self.EvFile.Filesets.Item(idx).DataPaths.Clear
#        return True
#
#
#    def addFilesetDataPathbyIndex(self, idx, datapath):
#        '''add the datapaths to a file set'''
#        if isinstance(datapath, Path):
#            datapath = str(datapath)
#        self.EvFile.Filesets.Item(idx).DataPaths.Insert(datapath, 2)
#        return True


    def clearDataPaths(self):
        '''Clear the datapaths'''
        self.EvFile.Properties.DataPaths.Clear()
        return True


    def addDataPaths(self, datapath):
        '''add the datapaths'''
        if isinstance(datapath, Path):
            datapath = str(datapath)
        self.EvFile.Properties.DataPaths.Insert(datapath, 2)
        return True


    def addDataFiles(self, fsdx, datafile):
        '''Add data files to the EV file'''
        if isinstance(datafile, Path):
            datafile = str(datafile)
        #self.EvFile.Filesets.Item(fsname).DataFiles.Add(datafile)
        self.EvFile.Filesets.Item(fsdx).DataFiles.Add(datafile)
        return True





class ProgressBar(QtWidgets.QWidget):
    def __init__(self):
        pass
    
    def MakeProgressBar(self, title, totalsize):
        self.title = title
        self.totalsize = totalsize
        self.progress = QtWidgets.QProgressDialog(self.title, "Cancel", 0, self.totalsize)
        #self.progress.setAutoClose(False)
        self.progress.setWindowModality(QtCore.Qt.WindowModal)
        self.progress.show()

    def HandleProgress(self, filenum):
        if self.totalsize > 0:
            process_percentage = filenum * 100/self.totalsize
            self.progress.setValue(process_percentage)
            #self.app.processEvents()

class Errors:
    def __init__(self):
        self.numerrors = 0
        #self.dir=QtWidgets.QFileDialog.getExistingDirectory(None, "Select Log File Output Directory")

    def createErrorFile(self, function, errorDir):
        '''Create the Error File in the specified directory
           Input: name of the file, directory name in which to create the file'''
        # check if the directory path is a pathlib object
        # if it is, use pathlib, if not convert to pathlib
        if not isinstance(errorDir, Path):
            errorDir = Path(errorDir)
        self.EFhndl = open(str((errorDir/Path('ErrorFile_'+function+'.csv'))), 'w')

    def appendErrorFile(self, message):
        '''Append an error message to the error file
           Input: error message'''
        self.EFhndl.write(message+'\n')
        self.numerrors += self.numerrors

    def closeErrorFile(self):
        self.EFhndl.close()

#class EVFile:
#    def __init__(self):
#        '''modules to operate on the EV file'''
#        pass
    

    
