Attribute VB_Name = "FunctionsModule"
Option Explicit 'This Module is specifically reserved for Functions

Declare PtrSafe Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long 'Used by GetUserName
Const NoError = 0 'The Function call was successful

Function AdjustUserName(UserName As String)
    
    If UserName = "penny.b.cummings" Then
        AdjustUserName = "Penny.A.Barton"
    Else
        AdjustUserName = UserName
    End If
    
End Function

Function GetEmailAddress(laboratory As String) As String
    
    Select Case laboratory
    
        Case "Albuquerque"
            GetEmailAddress = "DGXDALPROCESSING@questdiagnostics.com"
        Case "Atlanta"
            GetEmailAddress = ""
        Case "Auburn Hills"
            GetEmailAddress = ""
        Case "Baltimore"
            GetEmailAddress = ""
        Case "Cincinnati"
            GetEmailAddress = ""
        Case "Dallas"
            GetEmailAddress = "DGXDALPROCESSING@questdiagnostics.com"
        Case "Denver"
            GetEmailAddress = ""
        Case "DLO"
            GetEmailAddress = ""
        Case "Greensboro"
            GetEmailAddress = ""
        Case "Houston"
            GetEmailAddress = "DGXHOUPROCESSING@questdiagnostics.com"
        Case "Las Vegas"
            GetEmailAddress = ""
        Case "Lenexa"
            GetEmailAddress = ""
        Case "MACL"
            GetEmailAddress = ""
        Case "Marlborough"
            GetEmailAddress = ""
        Case "Miami"
            GetEmailAddress = "SpecimeninTransit_Re@questdiagnostics.com"
        Case "New Orleans"
            GetEmailAddress = ""
        Case "Philadelphia"
            GetEmailAddress = ""
        Case "Pittsburgh"
            GetEmailAddress = ""
        Case "Puerto Rico"
            GetEmailAddress = ""
        Case "Sacramento"
            GetEmailAddress = ""
        Case "Seattle"
            GetEmailAddress = ""
        Case "Solstas"
            GetEmailAddress = ""
        Case "Syosset"
            GetEmailAddress = ""
        Case "Tampa"
            GetEmailAddress = "DGXTPACNGREV@questdiagnostics.com"
        Case "Teterboro"
            GetEmailAddress = ""
        Case "Wallingford"
            GetEmailAddress = ""
                Case "West Hills"
            GetEmailAddress = ""
        Case "Wood Dale"
            GetEmailAddress = ""
        Case Else
            GetEmailAddress = ""
        
    End Select
    
End Function

Function GetUserName()

    Const lpnLength As Integer = 255 'Buffer size for the return string.
    
    Dim status As Integer 'Get return buffer space.
    Dim lpName, lpUserName As String 'For getting user information.

    lpUserName = Space$(lpnLength + 1) 'Assign the buffer size constant to lpUserName.
    status = WNetGetUser(lpName, lpUserName, lpnLength) 'Get the log-on name of the person using product.
    
    ' See whether error occurred.
    If status = NoError Then
       ' This line removes the null character. Strings in C are null-
       ' terminated. Strings in Visual Basic are not null-terminated.
       ' The null character must be removed from the C strings to be used
       ' cleanly in Visual Basic.
       lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
       MsgBox "Unable to get the name." 'An error occurred.
       End
    End If

    GetUserName = lpUserName 'Display the name of the person logged on to the machine.

End Function

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
