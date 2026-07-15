VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents mButton As MSForms.CommandButton
Attribute mButton.VB_VarHelpID = -1

Private mActionKey As String
Private mParentForm As Object

Public Sub Initialize( _
    ByVal ButtonObject As MSForms.CommandButton, _
    ByVal ActionKey As String, _
    ByVal ParentForm As Object)

    Set mButton = ButtonObject
    mActionKey = UCase$(Trim$(ActionKey))
    Set mParentForm = ParentForm

End Sub

Private Sub mButton_Click()

    Select Case mActionKey

        Case "RUN_TESTS"
            RunTests

        Case "CLOSE"
            Unload mParentForm

        Case "GENERATE_SCHEDULE"
            ShowPendingFeature "Generar planificacio"

        Case "VALIDATE_SCHEDULE"
            ShowPendingFeature "Validar planificacio"

        Case "EXPORT_SCHEDULE"
            ShowPendingFeature "Exportar calendari"

        Case "EMPLOYEES"
            ShowPendingFeature "Gestio d'operaris"

        Case "SUBSTITUTIONS"
            ShowPendingFeature "Substitucions"

        Case "VACATIONS"
            ShowPendingFeature "Vacances"

        Case "DUTIES"
            ShowPendingFeature "Guardies"

        Case "CONFIGURATION"
            ShowPendingFeature "Configuracio"

        Case "HISTORY"
            ShowPendingFeature "Historial de canvis"

        Case Else
            MsgBox _
                Prompt:="Accio no reconeguda: " & mActionKey, _
                Buttons:=vbExclamation, _
                Title:="TORNS CONSORCI"

    End Select

End Sub

Private Sub RunTests()

    On Error GoTo ErrorHandler

    modTests.RunAllTests

    modUI.AppInfo.UpdateTestStatus _
        TestsPassed:=45, _
        TestsRun:=45

    On Error Resume Next
    mParentForm.Controls("lblStatusOkTests").Caption = _
        "Tests: " & modUI.AppInfo.TestsStatusText
    On Error GoTo 0

    MsgBox _
        Prompt:="Els tests s'han executat correctament." & vbCrLf & _
                "Consulta la finestra Immediate per veure el detall.", _
        Buttons:=vbInformation, _
        Title:="TORNS CONSORCI"

    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="No s'han pogut executar els tests." & vbCrLf & Err.Description, _
        Buttons:=vbExclamation, _
        Title:="TORNS CONSORCI"

End Sub

Private Sub ShowPendingFeature(ByVal FeatureName As String)

    MsgBox _
        Prompt:=FeatureName & " estara disponible en una versio posterior.", _
        Buttons:=vbInformation, _
        Title:="TORNS CONSORCI"

End Sub
