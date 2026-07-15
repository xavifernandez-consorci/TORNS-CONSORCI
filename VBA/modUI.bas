Attribute VB_Name = "modUI"
Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modUI
' Layer:        Presentation
'
' Version:      1.0
'
' Purpose:
'   Centralize UserForm navigation and shared visual styling.
'
' Responsibilities:
'   - Maintain one shared clsTheme instance.
'   - Maintain one shared clsAppInfo instance.
'   - Open the main application form.
'   - Apply a consistent visual style to UserForms and controls.
'
' Restrictions:
'   - No business logic.
'   - No worksheet access.
'   - No persistence.
'   - No direct schedule generation.
'===============================================================================

Private Const MODULE_NAME As String = "modUI"

Private Const ERR_FORM_REQUIRED As Long = vbObjectError + 4000
Private Const ERR_MAIN_FORM As Long = vbObjectError + 4001

Private mTheme As clsTheme
Private mAppInfo As clsAppInfo

'===============================================================================
' Initializes the shared presentation services.
'===============================================================================
Public Sub InitializeUI()

    If mTheme Is Nothing Then
        Set mTheme = New clsTheme
    End If

    If mAppInfo Is Nothing Then
        Set mAppInfo = New clsAppInfo
    End If

End Sub

'===============================================================================
' Releases presentation-layer references.
'===============================================================================
Public Sub ReleaseUI()

    Set mAppInfo = Nothing
    Set mTheme = Nothing

End Sub

'===============================================================================
' Returns the shared application theme.
'===============================================================================
Public Function Theme() As clsTheme

    InitializeUI
    Set Theme = mTheme

End Function

'===============================================================================
' Returns the shared application information object.
'===============================================================================
Public Function AppInfo() As clsAppInfo

    InitializeUI
    Set AppInfo = mAppInfo

End Function

'===============================================================================
' Opens the main form.
'===============================================================================
Public Sub ShowMainForm()

    On Error GoTo ErrorHandler

    InitializeUI

    Load frmInici
    ApplyTheme frmInici

    frmInici.Show

    Exit Sub

ErrorHandler:
    Err.Raise _
        Number:=ERR_MAIN_FORM, _
        Source:=MODULE_NAME & ".ShowMainForm", _
        Description:="No s'ha pogut obrir el formulari principal. " & _
                    Err.Description

End Sub

'===============================================================================
' Applies the shared visual theme to one UserForm.
'
' The routine uses Object and TypeName so it can style standard VBA controls
' without adding additional library dependencies.
'===============================================================================
Public Sub ApplyTheme(ByVal FormObject As Object)

    Dim control As Object
    Dim currentTheme As clsTheme

    If FormObject Is Nothing Then
        RaiseUIError _
            ErrorNumber:=ERR_FORM_REQUIRED, _
            ProcedureName:="ApplyTheme", _
            Description:="El formulari no pot ser Nothing."
    End If

    Set currentTheme = Theme()

    With FormObject
        .BackColor = currentTheme.BackgroundColor
        .Font.Name = currentTheme.FontName
        .Font.Size = currentTheme.FontSize
    End With

    For Each control In FormObject.Controls
        ApplyControlTheme control, currentTheme
    Next control

End Sub

'===============================================================================
' Applies the appropriate style according to the control type.
'===============================================================================
Private Sub ApplyControlTheme( _
    ByVal ControlObject As Object, _
    ByVal CurrentTheme As clsTheme)

    ' Controls created by clsGrid already carry their own semantic colors.
    ' Do not overwrite them with the generic Label style.
    If IsGridControl(ControlObject) Then
        Exit Sub
    End If

    Select Case TypeName(ControlObject)

        Case "CommandButton"
            StyleCommandButton ControlObject, CurrentTheme

        Case "Label"
            StyleLabel ControlObject, CurrentTheme

        Case "Frame"
            StyleFrame ControlObject, CurrentTheme
            ApplyNestedControlsTheme ControlObject, CurrentTheme

        Case "TextBox", "ComboBox", "ListBox"
            StyleInputControl ControlObject, CurrentTheme

        Case "CheckBox", "OptionButton", "ToggleButton"
            StyleSelectionControl ControlObject, CurrentTheme

        Case "MultiPage", "Page"
            StyleContainer ControlObject, CurrentTheme
            ApplyNestedControlsTheme ControlObject, CurrentTheme

        Case Else
            ApplyDefaultFont ControlObject, CurrentTheme

    End Select

End Sub

'===============================================================================
' Applies the theme to controls contained inside frames, pages or multipages.
'===============================================================================
Private Sub ApplyNestedControlsTheme( _
    ByVal ContainerObject As Object, _
    ByVal CurrentTheme As clsTheme)

    Dim nestedControl As Object

    On Error GoTo CleanExit

    For Each nestedControl In ContainerObject.Controls
        ApplyControlTheme nestedControl, CurrentTheme
    Next nestedControl

CleanExit:
End Sub

'===============================================================================
' Returns True when the control belongs to a dynamically rendered grid.
'===============================================================================
Private Function IsGridControl(ByVal ControlObject As Object) As Boolean

    Dim controlTag As String

    On Error Resume Next
    controlTag = CStr(ControlObject.Tag)
    On Error GoTo 0

    IsGridControl = (Left$(controlTag, 5) = "GRID_")

End Function

'===============================================================================
' Standard button appearance.
'===============================================================================
Private Sub StyleCommandButton( _
    ByVal ButtonObject As Object, _
    ByVal CurrentTheme As clsTheme)

    With ButtonObject
        .BackColor = CurrentTheme.PrimaryColor
        .ForeColor = CurrentTheme.SurfaceColor
        .Font.Name = CurrentTheme.FontName
        .Font.Size = CurrentTheme.ButtonFontSize
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With

End Sub

'===============================================================================
' Standard label appearance.
'
' Special naming conventions:
'   lblTitle*   -> main title
'   lblSection* -> section heading
'   lblMuted*   -> secondary text
'   lblStatusOk* -> success text
'   lblStatusWarning* -> warning text
'   lblStatusError* -> error text
'===============================================================================
Private Sub StyleLabel( _
    ByVal LabelObject As Object, _
    ByVal CurrentTheme As clsTheme)

    Dim controlName As String

    controlName = LCase$(LabelObject.Name)

    With LabelObject
        .BackStyle = 0
        .ForeColor = CurrentTheme.TextColor
        .Font.Name = CurrentTheme.FontName
        .Font.Size = CurrentTheme.FontSize
        .Font.Bold = False
    End With

    If Left$(controlName, 8) = "lbltitle" Then

        With LabelObject
            .ForeColor = CurrentTheme.PrimaryDarkColor
            .Font.Size = CurrentTheme.TitleFontSize
            .Font.Bold = True
        End With

    ElseIf Left$(controlName, 10) = "lblsection" Then

        With LabelObject
            .ForeColor = CurrentTheme.PrimaryColor
            .Font.Size = CurrentTheme.SectionFontSize
            .Font.Bold = True
        End With

    ElseIf Left$(controlName, 8) = "lblmuted" Then

        LabelObject.ForeColor = CurrentTheme.MutedTextColor

    ElseIf Left$(controlName, 11) = "lblstatusok" Then

        LabelObject.ForeColor = CurrentTheme.SuccessColor

    ElseIf Left$(controlName, 16) = "lblstatuswarning" Then

        LabelObject.ForeColor = CurrentTheme.WarningColor

    ElseIf Left$(controlName, 14) = "lblstatuserror" Then

        LabelObject.ForeColor = CurrentTheme.ErrorColor

    End If

End Sub

'===============================================================================
' Standard frame appearance.
'===============================================================================
Private Sub StyleFrame( _
    ByVal FrameObject As Object, _
    ByVal CurrentTheme As clsTheme)

    With FrameObject
        .BackColor = CurrentTheme.SurfaceColor
        .ForeColor = CurrentTheme.PrimaryColor
        .Font.Name = CurrentTheme.FontName
        .Font.Size = CurrentTheme.SectionFontSize
        .Font.Bold = True
    End With

End Sub

'===============================================================================
' Standard input-control appearance.
'===============================================================================
Private Sub StyleInputControl( _
    ByVal InputObject As Object, _
    ByVal CurrentTheme As clsTheme)

    With InputObject
        .BackColor = CurrentTheme.SurfaceColor
        .ForeColor = CurrentTheme.TextColor
        .Font.Name = CurrentTheme.FontName
        .Font.Size = CurrentTheme.FontSize
    End With

End Sub

'===============================================================================
' Standard checkbox, option-button and toggle-button appearance.
'===============================================================================
Private Sub StyleSelectionControl( _
    ByVal SelectionObject As Object, _
    ByVal CurrentTheme As clsTheme)

    With SelectionObject
        .BackColor = CurrentTheme.BackgroundColor
        .ForeColor = CurrentTheme.TextColor
        .Font.Name = CurrentTheme.FontName
        .Font.Size = CurrentTheme.FontSize
    End With

End Sub

'===============================================================================
' Standard container appearance.
'===============================================================================
Private Sub StyleContainer( _
    ByVal ContainerObject As Object, _
    ByVal CurrentTheme As clsTheme)

    On Error Resume Next
    ContainerObject.BackColor = CurrentTheme.BackgroundColor
    ContainerObject.ForeColor = CurrentTheme.TextColor
    ContainerObject.Font.Name = CurrentTheme.FontName
    ContainerObject.Font.Size = CurrentTheme.FontSize
    On Error GoTo 0

End Sub

'===============================================================================
' Applies only typography to unsupported control types.
'===============================================================================
Private Sub ApplyDefaultFont( _
    ByVal ControlObject As Object, _
    ByVal CurrentTheme As clsTheme)

    On Error Resume Next
    ControlObject.Font.Name = CurrentTheme.FontName
    ControlObject.Font.Size = CurrentTheme.FontSize
    On Error GoTo 0

End Sub

'===============================================================================
' Raises a controlled presentation-layer error.
'===============================================================================
Private Sub RaiseUIError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=MODULE_NAME & "." & ProcedureName, _
        Description:=Description

End Sub
