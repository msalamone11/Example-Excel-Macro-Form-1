Attribute VB_Name = "SubRoutinesModule"
Option Explicit

Sub DeleteButtons()

    Sheets("Specimen In Transit Form").Select
    ActiveSheet.Shapes.Range(Array("GetFromQlsButton")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("ResetFormButton")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("SendFormButton")).Select
    Selection.Delete
    
End Sub

Sub FillInCsrName()

    Range("CsrName") = Application.WorksheetFunction.Proper(GetUserName)

End Sub

Sub FillInDate()

    Range("Date").Value = Format(Now(), "mm/dd/yyyy")

End Sub

Sub GetFromQls()

    Dim UserName As String
        
    UserName = GetUserName
    
    UserName = AdjustUserName(UserName) 'Some people's names in their C:\Users\%UserProfile% do not reflect what is returned by the GetUserName function.
        
    OpenAnyFile ("C:\Users\" & UserName & "\NCS-Automated-Forms\Specimen-In-Transit-Screen-Scraper\Specimen-In-Transit-Screen-Scraper.exe")

End Sub

Sub OpenAnyFile(FilePath As String)

    Dim fileX As Object
    
    Set fileX = CreateObject("Shell.Application")
    
    fileX.Open (FilePath)

End Sub

Sub PopulateFormDataToLog(macroName As String)

        On Error Resume Next

        Dim FILE_NAME As String
        FILE_NAME = "\\QDCNS0002\TMP_Data$\TMPDept\Knowledge_Base\Logs\Specimen-In-Transit-Form\Specimen-In-Transit-Form-Log.txt"
        Dim strFileExists As String
        strFileExists = Dir(FILE_NAME) 'Dir(FILE_NAME) will return the file name if it exists.
        
        If strFileExists = "Specimen-In-Transit-Form-Log.txt" Then
        
            Dim csrName As String
            Dim callersName As String
            Dim accountName As String
            Dim accountNumber As String
            Dim AccessionNumber As String
            Dim reqNumber As String
            Dim patientsName As String
            Dim patientsDob As String
            
            Dim laboratory As String
            Dim routine As String
            Dim stat As String
            Dim specialHandle As String
            Dim addTests As String
            Dim cancelTests As String
            
            Dim testName1 As String
            Dim testName2 As String
            Dim testName3 As String
            Dim testName4 As String
            Dim testCode1 As String
            Dim testCode2 As String
            Dim testCode3 As String
            Dim testCode4 As String
            
            Dim specialHandlingInstructions As String
            Dim transportationMethodAndEta As String
            
            csrName = GetUserName
            callersName = Range("CallersName")
            accountName = Range("AccountName")
            accountNumber = Range("AccountNumber")
            AccessionNumber = Range("AccessionNumber")
            reqNumber = Range("ReqNumber")
            patientsName = Range("PatientsName")
            patientsDob = Range("PatientsDob")
            
            laboratory = Range("laboratory")
            routine = Range("routine")
            stat = Range("stat")
            addTests = Range("addTests")
            cancelTests = Range("cancelTests")
            
            testName1 = Range("TestName1")
            testName2 = Range("TestName2")
            testName3 = Range("TestName3")
            testName4 = Range("TestName4")
            testCode1 = Range("TestCode1")
            testCode2 = Range("TestCode2")
            testCode3 = Range("TestCode3")
            testCode4 = Range("TestCode4")
            
            specialHandlingInstructions = Range("specialHandlingInstructions")
            transportationMethodAndEta = Range("transportationMethodAndEta")
        
            Dim inputString As String

            inputString = Now & "|" & csrName & "|" & macroName & "|" & callersName & "|" & accountName & "|" & accountNumber & "|" & AccessionNumber _
                          & "|" & reqNumber & "|" & patientsName & "|" & patientsDob & "|" & laboratory & "|" & routine & "|" & stat _
                          & "|" & addTests & "|" & cancelTests & "|" & testName1 & "|" & testCode1 & "|" & testName2 & "|" & testCode2 _
                          & "|" & testName3 & "|" & testCode3 & "|" & testName4 & "|" & testCode4 _
                          & "|" & specialHandlingInstructions & "|" & transportationMethodAndEta
            
            Open FILE_NAME For Append As #1
            Write #1, inputString
            Close #1

        End If

    End Sub

Sub ResetForm()

    Application.ScreenUpdating = False

    Dim RangesToClear() As Variant
    Dim Item As Variant
    
    RangesToClear = Array("Date", "CsrName", "CallersName", _
                                            "AccountName", "AccountNumber", "AccessionNumber", _
                                            "ReqNumber", "PatientsName", "PatientsDob", _
                                            "Laboratory", "Routine", "Stat", _
                                            "SpecialHandle", "AddTests", "CancelTests", _
                                            "TestName1", "TestName2", "TestName3", _
                                            "TestName4", "TestCode1", "TestCode2", _
                                            "TestCode3", "TestCode4", "SpecialHandlingInstructions", _
                                            "TransportationMethodAndEta")
    
    For Each Item In RangesToClear
        
        Range(Item) = ""
        
    Next
    
    Call FillInDate
    Call FillInCsrName
    
    Application.Goto Worksheets("Specimen In Transit Form").Range("CallersName"), False
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Range("CallersName").Select
    
    Application.ScreenUpdating = True

End Sub

Sub SendForm()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False  'or True
      
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim AccessionNumber As String
    Dim Body As String
    Dim DFax, DEmail, SEmail, Subj As String
    Dim sBody As Range
    Dim rng As Range
    Dim emailAddress As String
        
    Worksheets("Specimen In Transit Form").Unprotect Password:=""
    
    emailAddress = GetEmailAddress(Range("Laboratory").Value)
    
    Set rng = Nothing
    Set rng = Range("EntireForm").SpecialCells(xlCellTypeVisible)
    
    Range("Date").Value = Format(Now(), "mm/dd/yyyy")
    Range("CsrName") = Application.WorksheetFunction.Proper(GetUserName)
    
    Set sBody = Nothing
    
    Sheets("Specimen In Transit Form").Select
    
    AccessionNumber = Range("AccessionNumber").Value
    
    Subj = AccessionNumber & " - Specimen In Transit Form"
    
    Sheets("Specimen In Transit Form").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    
    Set Sourcewb = ActiveWorkbook
    
    'Copy the sheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook
    
    'Call DeleteExtraText
    Call DeleteButtons

    'Set Sourcewb = ActiveWorkbook
    '
    '    'Copy the sheet to a new workbook
    '    ActiveSheet.Copy
    '    Set Destwb = ActiveWorkbook
    
    'Determine the Excel version and file extension/format
    With Destwb
    If Val(Application.Version) < 12 Then
        'You use Excel 2000-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        FileExtStr = ".xlsx"
        FileFormatNum = 51
    
    End If
    End With

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = AccessionNumber & " - Specimen In Transit Form"
        
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, _
                FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = emailAddress
            .CC = ""
            .BCC = ""
            .Subject = Subj
            .Attachments.Add Destwb.FullName 'You can add other files also like this: .Attachments.Add ("C:\test.txt")
            .Display 'Or .Send
            .HTMLBody = RangetoHTML(rng)
        End With
    End With
   
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Destwb.Close False
    
    Sourcewb.Activate
    
    Range("D9").Select
    Call ResetForm
    
    Worksheets("Specimen In Transit Form").Protect Password:=""
    
    'Application.Quit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ValidateBlankFormRanges(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        
        If Range(Item) = "" Then
        
            Range(Item).Select
            MsgBox ("You've missed a required field. Please fill in this field and then try again.")
            
            Call PopulateFormDataToLog("ValidateBlankFormRanges")
            Debug.Print ("Caught a blank form field: " & Item)
            
            End
        
        End If
        
    Next
    
    'Example Of CommaSeparatedStringOfRanges would be "'Range1','Range2','Range3'"

End Sub

Sub ValidateSpecimenInTransitForm()

    Call ValidateBlankFormRanges("Date,CsrName,CallersName,AccountName,AccountNumber,AccessionNumber,ReqNumber," & _
                                 "PatientsName,PatientsDob,Laboratory,TestName1,TestCode1")
    
End Sub
