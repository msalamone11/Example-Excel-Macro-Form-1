VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub GetFromQlsButton_Click()

    Call GetFromQls
    Call PopulateFormDataToLog("GetFromQls")

End Sub

Private Sub ResetFormButton_Click()

    Call PopulateFormDataToLog("ResetForm")
    Call ResetForm

End Sub

Private Sub SendFormButton_Click()

    Call ValidateSpecimenInTransitForm
    Call PopulateFormDataToLog("SendForm")
    Call SendForm

End Sub
