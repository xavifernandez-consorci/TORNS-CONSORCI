Attribute VB_Name = "modUIBuilder"
Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modUIBuilder
' Layer:        Presentation
'
' Version:      1.1
'
' Purpose:
'   Build frmInici dynamically at runtime.
'
' Responsibilities:
'   - Build the side navigation menu.
'   - Build the system-status panel.
'   - Build the current-week dashboard.
'   - Render the reusable weekly grid.
'
' Notes:
'   - frmInici only needs to exist as an empty UserForm.
'   - Controls are recreated every time the form is initialized.
'   - The weekly grid currently uses demonstration data.
'===============================================================================

Private mMainFormHandlers As Collection
Private mWeeklyGrid As clsGrid

'===============================================================================
' Builds the complete main dashboard.
'===============================================================================
Public Sub BuildMainForm(ByVal FormObject As Object)

    Dim info As clsAppInfo
    Dim frameMenu As MSForms.Frame
    Dim frameStatus As MSForms.Frame
    Dim frameWeek As MSForms.Frame
    Dim weekMonday As Date

    Set info = modUI.AppInfo()
    Set mMainFormHandlers = New Collection

    ClearFormControls FormObject
    ConfigureMainForm FormObject, info

    AddLabel _
        ParentObject:=FormObject, _
        ControlName:="lblTitleMain", _
        CaptionText:=info.Name, _
        LeftValue:=24, _
        TopValue:=16, _
        WidthValue:=520, _
        HeightValue:=30, _
        StyleKey:="TITLE"

    AddLabel _
        ParentObject:=FormObject, _
        ControlName:="lblMutedSubtitle", _
        CaptionText:=info.Description, _
        LeftValue:=24, _
        TopValue:=50, _
        WidthValue:=620, _
        HeightValue:=20, _
        StyleKey:="MUTED"

    AddLabel _
        ParentObject:=FormObject, _
        ControlName:="lblMutedVersion", _
        CaptionText:="Versio " & info.VersionText, _
        LeftValue:=820, _
        TopValue:=24, _
        WidthValue:=150, _
        HeightValue:=18, _
        StyleKey:="MUTED"

    Set frameMenu = AddFrame( _
        ParentObject:=FormObject, _
        ControlName:="fraMenu", _
        CaptionText:="MENU", _
        LeftValue:=24, _
        TopValue:=88, _
        WidthValue:=170, _
        HeightValue:=390)

    BuildSideMenu frameMenu, FormObject

    Set frameStatus = AddFrame( _
        ParentObject:=FormObject, _
        ControlName:="fraEstat", _
        CaptionText:="ESTAT", _
        LeftValue:=24, _
        TopValue:=492, _
        WidthValue:=170, _
        HeightValue:=150)

    BuildStatusPanel frameStatus, info

    Set frameWeek = AddFrame( _
        ParentObject:=FormObject, _
        ControlName:="fraSetmana", _
        CaptionText:="SETMANA EN CURS", _
        LeftValue:=210, _
        TopValue:=88, _
        WidthValue:=760, _
        HeightValue:=554)

    weekMonday = CurrentWeekMonday()

    AddLabel _
        ParentObject:=frameWeek, _
        ControlName:="lblSectionWeek", _
        CaptionText:=FormatWeekRange(weekMonday), _
        LeftValue:=14, _
        TopValue:=24, _
        WidthValue:=360, _
        HeightValue:=20, _
        StyleKey:="SECTION"

    AddLabel _
        ParentObject:=frameWeek, _
        ControlName:="lblMutedGridNotice", _
        CaptionText:="Dades de demostracio. La connexio amb la planificacio real es fara al proper pas.", _
        LeftValue:=390, _
        TopValue:=24, _
        WidthValue:=345, _
        HeightValue:=20, _
        StyleKey:="MUTED"

    BuildWeeklyGrid frameWeek, weekMonday

    AddLegend frameWeek

    modUI.ApplyTheme FormObject

End Sub

'===============================================================================
' Releases runtime-created objects.
'===============================================================================
Public Sub ReleaseMainForm()

    Set mWeeklyGrid = Nothing
    Set mMainFormHandlers = Nothing

End Sub

'===============================================================================
' Side menu
'===============================================================================
Private Sub BuildSideMenu( _
    ByVal MenuFrame As MSForms.Frame, _
    ByVal MainForm As Object)

    AddButton MenuFrame, "cmdGenerarPlanificacio", _
        "Generar", 14, 30, 140, 38, "GENERATE_SCHEDULE", MainForm

    AddButton MenuFrame, "cmdOperaris", _
        "Operaris", 14, 76, 140, 38, "EMPLOYEES", MainForm

    AddButton MenuFrame, "cmdSubstitucions", _
        "Substitucions", 14, 122, 140, 38, "SUBSTITUTIONS", MainForm

    AddButton MenuFrame, "cmdVacances", _
        "Vacances", 14, 168, 140, 38, "VACATIONS", MainForm

    AddButton MenuFrame, "cmdGuardies", _
        "Guardies", 14, 214, 140, 38, "DUTIES", MainForm

    AddButton MenuFrame, "cmdConfiguracio", _
        "Configuracio", 14, 260, 140, 38, "CONFIGURATION", MainForm

    AddButton MenuFrame, "cmdExecutarTests", _
        "Executar tests", 14, 306, 140, 38, "RUN_TESTS", MainForm

    AddButton MenuFrame, "cmdTancar", _
        "Tancar", 14, 352, 140, 28, "CLOSE", MainForm

End Sub

'===============================================================================
' Status panel
'===============================================================================
Private Sub BuildStatusPanel( _
    ByVal StatusFrame As MSForms.Frame, _
    ByVal Info As clsAppInfo)

    AddLabel StatusFrame, "lblStatusOkOperaris", _
        "Operaris actius: 14", 14, 28, 140, 18, "STATUS_OK"

    AddLabel StatusFrame, "lblStatusWarningBaixes", _
        "Baixes: 1", 14, 52, 140, 18, "STATUS_WARNING"

    AddLabel StatusFrame, "lblStatusWarningVacances", _
        "Vacances: 1", 14, 76, 140, 18, "STATUS_WARNING"

    AddLabel StatusFrame, "lblStatusOkTests", _
        "Tests: " & Info.TestsStatusText, 14, 100, 140, 18, "STATUS_OK"

    AddLabel StatusFrame, "lblMutedEnvironment", _
        Info.EnvironmentName, 14, 124, 140, 18, "MUTED"

End Sub

'===============================================================================
' Weekly grid
'===============================================================================
Private Sub BuildWeeklyGrid( _
    ByVal WeekFrame As MSForms.Frame, _
    ByVal WeekMonday As Date)

    Set mWeeklyGrid = New clsGrid

    mWeeklyGrid.Initialize _
        ContainerObject:=WeekFrame, _
        GridName:="CurrentWeek", _
        LeftValue:=14, _
        TopValue:=54

    mWeeklyGrid.SetDimensions _
        NameWidth:=130, _
        StatusWidth:=78, _
        CellWidth:=68, _
        RowHeight:=24, _
        HeaderHeight:=34

    mWeeklyGrid.SetWeekHeaders WeekMonday

    AddDemonstrationRows mWeeklyGrid
    mWeeklyGrid.Render

End Sub

'===============================================================================
' Demonstration data for the first dashboard version.
'===============================================================================
Private Sub AddDemonstrationRows(ByVal Grid As clsGrid)

    Grid.AddRow "Operari 01", "Actiu", _
        Array("D", "D", "D", "I", "I", "I", "I")

    Grid.AddRow "Operari 02", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 03", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 04", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 05", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 06", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 07", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 08", "Actiu", _
        Array("M", "M", "M", "M", "M", "M", "M")

    Grid.AddRow "Operari 09", "Actiu", _
        Array("T", "T", "T", "T", "T", "T", "T")

    Grid.AddRow "Operari 10", "Actiu", _
        Array("T", "T", "T", "T", "T", "T", "T")

    Grid.AddRow "Operari 11", "Actiu", _
        Array("T", "T", "T", "T", "T", "T", "T")

    Grid.AddRow "Operari 12", "Actiu", _
        Array("T", "T", "T", "T", "T", "T", "T")

    Grid.AddRow "Operari 13", "Baixa", _
        Array("B", "B", "B", "B", "B", "B", "B")

    Grid.AddRow "Operari 14", "Vacances", _
        Array("V", "V", "V", "V", "V", "V", "V")

End Sub

'===============================================================================
' Legend
'===============================================================================
Private Sub AddLegend(ByVal ParentObject As Object)

    Dim labels As Variant
    Dim codes As Variant
    Dim index As Long
    Dim currentLeft As Single

    codes = Array("M", "T", "I", "D", "B", "V")
    labels = Array("Mati", "Tarda", "Intensiu", "Descans", "Baixa", "Vacances")

    currentLeft = 14

    For index = LBound(codes) To UBound(codes)

        AddLegendItem _
            ParentObject:=ParentObject, _
            ShiftCode:=CStr(codes(index)), _
            Description:=CStr(labels(index)), _
            LeftValue:=currentLeft, _
            TopValue:=500

        currentLeft = currentLeft + 116

    Next index

End Sub

Private Sub AddLegendItem( _
    ByVal ParentObject As Object, _
    ByVal ShiftCode As String, _
    ByVal Description As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single)

    Dim codeLabel As MSForms.Label
    Dim descriptionLabel As MSForms.Label

    Set codeLabel = ParentObject.Controls.Add( _
        "Forms.Label.1", _
        "lblLegendCode" & ShiftCode, _
        True)

    With codeLabel
        .Caption = ShiftCode
        .Left = LeftValue
        .Top = TopValue
        .Width = 24
        .Height = 20
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .Font.Bold = True
        .BackStyle = fmBackStyleOpaque
        .BackColor = LegendBackColor(ShiftCode)
        .ForeColor = LegendTextColor(ShiftCode)
    End With

    Set descriptionLabel = ParentObject.Controls.Add( _
        "Forms.Label.1", _
        "lblLegendText" & ShiftCode, _
        True)

    With descriptionLabel
        .Caption = Description
        .Left = LeftValue + 28
        .Top = TopValue + 2
        .Width = 82
        .Height = 18
        .BackStyle = fmBackStyleTransparent
        .ForeColor = modUI.Theme.TextColor
    End With

End Sub

'===============================================================================
' Main-form configuration
'===============================================================================
Private Sub ConfigureMainForm( _
    ByVal FormObject As Object, _
    ByVal Info As clsAppInfo)

    With FormObject
        .Caption = Info.DisplayTitle
        .Width = 1000
        .Height = 700
        .StartUpPosition = 2
        .ScrollBars = 0
        .KeepScrollBarsVisible = 0
    End With

End Sub

Private Sub ClearFormControls(ByVal FormObject As Object)

    Dim index As Long

    For index = FormObject.Controls.Count - 1 To 0 Step -1
        FormObject.Controls.Remove FormObject.Controls(index).Name
    Next index

End Sub

'===============================================================================
' Generic control creation
'===============================================================================
Private Function AddFrame( _
    ByVal ParentObject As Object, _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single, _
    ByVal WidthValue As Single, _
    ByVal HeightValue As Single) As MSForms.Frame

    Dim frame As MSForms.Frame

    Set frame = ParentObject.Controls.Add( _
        "Forms.Frame.1", _
        ControlName, _
        True)

    With frame
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = WidthValue
        .Height = HeightValue
    End With

    Set AddFrame = frame

End Function

Private Sub AddLabel( _
    ByVal ParentObject As Object, _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single, _
    ByVal WidthValue As Single, _
    ByVal HeightValue As Single, _
    ByVal StyleKey As String)

    Dim label As MSForms.Label

    Set label = ParentObject.Controls.Add( _
        "Forms.Label.1", _
        ControlName, _
        True)

    With label
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = WidthValue
        .Height = HeightValue
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .Font.Name = modUI.Theme.FontName
        .Font.Size = modUI.Theme.FontSize
    End With

    Select Case UCase$(StyleKey)

        Case "TITLE"
            label.Font.Size = modUI.Theme.TitleFontSize
            label.Font.Bold = True
            label.ForeColor = modUI.Theme.PrimaryDarkColor

        Case "SECTION"
            label.Font.Size = modUI.Theme.SectionFontSize
            label.Font.Bold = True
            label.ForeColor = modUI.Theme.PrimaryColor

        Case "STATUS_OK"
            label.Font.Bold = True
            label.ForeColor = modUI.Theme.SuccessColor

        Case "STATUS_WARNING"
            label.Font.Bold = True
            label.ForeColor = modUI.Theme.WarningColor

        Case Else
            label.ForeColor = modUI.Theme.MutedTextColor

    End Select

End Sub

Private Sub AddButton( _
    ByVal ParentObject As Object, _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single, _
    ByVal WidthValue As Single, _
    ByVal HeightValue As Single, _
    ByVal ActionKey As String, _
    ByVal MainForm As Object)

    Dim button As MSForms.CommandButton
    Dim handler As clsMenuButtonHandler

    Set button = ParentObject.Controls.Add( _
        "Forms.CommandButton.1", _
        ControlName, _
        True)

    With button
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = WidthValue
        .Height = HeightValue
        .Tag = ActionKey
    End With

    Set handler = New clsMenuButtonHandler

    handler.Initialize _
        ButtonObject:=button, _
        ActionKey:=ActionKey, _
        ParentForm:=MainForm

    mMainFormHandlers.Add handler

End Sub

'===============================================================================
' Week helpers
'===============================================================================
Private Function CurrentWeekMonday() As Date

    CurrentWeekMonday = _
        DateAdd( _
            "d", _
            1 - Weekday(Date, vbMonday), _
            Date)

End Function

Private Function FormatWeekRange(ByVal WeekMonday As Date) As String

    FormatWeekRange = _
        Format$(WeekMonday, "dd/mm/yyyy") & _
        " - " & _
        Format$(DateAdd("d", 6, WeekMonday), "dd/mm/yyyy")

End Function

'===============================================================================
' Legend color helpers
'===============================================================================
Private Function LegendBackColor(ByVal ShiftCode As String) As Long

    Select Case UCase$(ShiftCode)

        Case "M"
            LegendBackColor = RGB(189, 215, 238)

        Case "T"
            LegendBackColor = RGB(248, 203, 173)

        Case "I"
            LegendBackColor = modUI.Theme.PrimaryColor

        Case "D"
            LegendBackColor = RGB(217, 217, 217)

        Case "B"
            LegendBackColor = RGB(244, 177, 177)

        Case "V"
            LegendBackColor = RGB(198, 224, 180)

        Case Else
            LegendBackColor = modUI.Theme.SurfaceColor

    End Select

End Function

Private Function LegendTextColor(ByVal ShiftCode As String) As Long

    If UCase$(ShiftCode) = "I" Then
        LegendTextColor = modUI.Theme.SurfaceColor
    Else
        LegendTextColor = modUI.Theme.TextColor
    End If

End Function
