Attribute VB_Name = "Step_0_Start"
'---------------------------------------------------------------------
' Date Created : May 15, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : May 15, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Start_Here
' Description  : The purpose of function is to initialize the userform.
'---------------------------------------------------------------------
Sub Start_Here()
   
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String, strLabel3 As String
    Dim frameLabel1 As String, frameLabel2 As String, frameLabel3 As String
    Dim userFormCaption As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Label Strings
    userFormCaption = "KIENZLE LAB TOOLS"
    button1 = "CREATE NEW RCM COMPOSITE FILES"
    frameLabel2 = "TOOL GUIDE"
    frameLabel3 = "HELP SECTION"
    
    strLabel1 = "THE COMPLETE RCM COMPOSITE FILE MACRO"
    strLabel2 = "For more information, hover mouse over button."
    
    ' UserForm Initialize
    RCM_COMPF7_Form.Caption = userFormCaption
    RCM_COMPF7_Form.Frame2.Caption = frameLabel2
    RCM_COMPF7_Form.Frame5.Caption = frameLabel3
    RCM_COMPF7_Form.Frame2.Font.Bold = True
    RCM_COMPF7_Form.Frame5.Font.Bold = True
    RCM_COMPF7_Form.Label1.Caption = strLabel1
    RCM_COMPF7_Form.Label1.Font.Size = 21
    RCM_COMPF7_Form.Label1.Font.Bold = True
    RCM_COMPF7_Form.Label1.TextAlign = fmTextAlignCenter
    
    RCM_COMPF7_Form.CommandButton1.Caption = button1
    RCM_COMPF7_Form.CommandButton1.Font.Size = 11
    
    ' Help File
    RCM_COMPF7_Form.Label2 = strLabel2
    RCM_COMPF7_Form.Label2.Font.Size = 8
    RCM_COMPF7_Form.Label2.Font.Italic = True
    
    Application.StatusBar = "Macro has been initiated."
    RCM_COMPF7_Form.Show

End Sub
'---------------------------------------------------------------------------------------
' Date Created : July 18, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HELPFILE
' Description  : This function will feed the help tip section depending on the button
'                that has been activated.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function HELPFILE(ByVal Notification As Integer) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case 1
            NotifyUser = "TITLE: COMPLETE RCM COMPOSITE FILE MACRO" & vbLf
            NotifyUser = NotifyUser & "DESCRIPTION: This macro will append " & _
                "re-create the original composite files by using the RCM values. " & vbLf
            NotifyUser = NotifyUser & "INPUT: Find the location of the folder containing all RCM grid files." & vbLf
            NotifyUser = NotifyUser & "OUTPUT: New composite .TXT files" & vbLf
    End Select
    
    HELPFILE = NotifyUser
    
End Function
