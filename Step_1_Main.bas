Attribute VB_Name = "Step_1_Main"
Public objFSOlog As Object
Public logfile As TextStream
Public logtxt As String
Public appSTATUS As String
'---------------------------------------------------------------------------------------
' Date Created : May 15, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 3, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RCM_COMPF7_MAIN
' Description  : This is the main function that ties in two other sub-main functions.
'                First, this function sets up the folders and validates the user input.
'                It then calls the ACRU_COMPF7_ProcessingZonalStat function to process
'                the .DBF files. After all the .OUT files have been created, it calls
'                on the ACRU_COMPF7_CompositeFile to create a new AB10K grid file and
'                a new composite file, of which both contains 7 variables.
'---------------------------------------------------------------------------------------
Function RCM_COMPF7_MAIN()

    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String
    
    Dim UserSelectedFolder As String
    Dim MAINFolder As String
    Dim CFDIR As String
    Dim CopiedFiles As Long
    Dim MainOUT As String
    Dim CFOUT As String
    Dim ResultCF As Boolean
    Dim subARRAY() As String, outARRAY() As String
    Dim refIDArray() As String
    Dim refIndex As Integer
    
    ' Initialize Variables
    SummaryTitle = "Diagnostic Summary"

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    start_time = Now()
    
    '---------------------------------------------------------------------
    ' I. USER INPUT
    '---------------------------------------------------------------------
    UserSelectedFolder = GetFolder
    Debug.Print UserSelectedFolder
    If Len(UserSelectedFolder) = 0 Then GoTo Cancel
    MAINFolder = ReturnFolderName(UserSelectedFolder)
    Debug.Print MAINFolder

    '---------------------------------------------------------------------
    ' II. SETUP LOGFILE
    '---------------------------------------------------------------------
    CFDIR = MAINFolder & "_" & "CFOUT"
    Call CreateNewFolder(UserSelectedFolder, CFDIR)    ' Create the Composite File Directory
    CFOUT = ReturnSubFolder(UserSelectedFolder, CFDIR)
    
    ' Setup Log File
    Dim logfilename As String, logtextfile As String
    logfilename = MAINFolder & "rcm_log.txt"
    logtextfile = CFOUT & logfilename
    If Right(CFOUT, 1) <> "\" Then logtextfile = CFOUT & "\" & logfilename

    Set objFSOlog = CreateObject("Scripting.FileSystemObject")
    Set logfile = objFSOlog.CreateTextFile(logtextfile, True)
    
	'---------------------------------------------------------------------
    ' III. Maintain log starting from here
	'---------------------------------------------------------------------
    logfile.WriteLine " [ Start of Program. ] "
    logfile.WriteLine "Selected directory: " & UserSelectedFolder
    logfile.WriteLine "Main directory: " & MAINFolder
    logfile.WriteLine "Output directory: " & CFOUT
    
	ResultCF = RCM_COMPF7_CompositeFile(MAINFolder, CFOUT)
    
    logfile.WriteLine " [ End of Program. ] "
	
    '---------------------------------------------------------------------
	' IV. Total Program Processing Time
    '---------------------------------------------------------------------
    end_time = Now()
    ProcessingTime = DateDiff("n", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime)
    MsgBox MessageSummary, vbOKOnly, SummaryTitle

	'---------------------------------------------------------------------
    ' V. Close Log File
	'---------------------------------------------------------------------
    logfile.Close
    Set logfile = Nothing
    Set objFSOlog = Nothing
    
Cancel:
    If Len(UserSelectedFolder) = 0 Then
        MsgBox "No folder selected.", vbOKOnly, SummaryTitle
    End If
End Function

'---------------------------------------------------------------------------------------
' Date Created : May 15, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 3, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ACRU_COMPF7_CompositeFile
' Description  : This function will process the old AB10K grid files and .OUT files
'                in order to create the new composite file which contains 7 variables.
'---------------------------------------------------------------------------------------
Function RCM_COMPF7_CompositeFile(ByVal sourceDIR As String, ByVal outDIR As String) As Boolean

    Dim Result As Boolean
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    appSTATUS = "In progress: Creating new composite files..."
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    
    Result = ProcessCompositeFiles(sourceDIR, outDIR)
	
    If Result = False Then RCM_COMPF7_CompositeFile = False
    If Result = True Then RCM_COMPF7_CompositeFile = True
    
    Application.StatusBar = False

End Function
