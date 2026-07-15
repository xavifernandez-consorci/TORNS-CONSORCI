VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    clsGrid
' Layer:        Presentation
'
' Version:      1.0
'
' Purpose:
'   Render a reusable weekly grid inside any UserForm container.
'
' Responsibilities:
'   - Store column headers and row data.
'   - Render operator names, status and daily shift cells.
'   - Apply consistent colors according to the shift code.
'   - Clear and redraw only the controls created by this grid.
'
' Restrictions:
'   - No business logic.
'   - No schedule generation.
'   - No worksheet access.
'   - No persistence.
'===============================================================================

Private Const GRID_CONTROL_PREFIX As String = "GRID_"

Private Const DEFAULT_ROW_HEIGHT As Single = 24
Private Const DEFAULT_HEADER_HEIGHT As Single = 28
Private Const DEFAULT_NAME_WIDTH As Single = 135
Private Const DEFAULT_STATUS_WIDTH As Single = 95
Private Const DEFAULT_CELL_WIDTH As Single = 46
Private Const DEFAULT_GAP As Single = 2

Private mContainer As Object
Private mTheme As clsTheme
Private mGridName As String

Private mHeaders As Collection
Private mRowNames As Collection
Private mRowStates As Collection
Private mRows As Collection

Private mLeft As Single
Private mTop As Single
Private mNameWidth As Single
Private mStatusWidth As Single
Private mCellWidth As Single
Private mRowHeight As Single
Private mHeaderHeight As Single

Private Sub Class_Initialize()

    Set mHeaders = New Collection
    Set mRowNames = New Collection
    Set mRowStates = New Collection
    Set mRows = New Collection

    mLeft = 0
    mTop = 0
    mNameWidth = DEFAULT_NAME_WIDTH
    mStatusWidth = DEFAULT_STATUS_WIDTH
    mCellWidth = DEFAULT_CELL_WIDTH
    mRowHeight = DEFAULT_ROW_HEIGHT
    mHeaderHeight = DEFAULT_HEADER_HEIGHT

End Sub

'===============================================================================
' Connects the grid to a UserForm, Frame, Page or other controls container.
'===============================================================================
Public Sub Initialize( _
    ByVal ContainerObject As Object, _
    ByVal GridName As String, _
    Optional ByVal LeftValue As Single = 0, _
    Optional ByVal TopValue As Single = 0)

    If ContainerObject Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 5000, _
            Source:="clsGrid.Initialize", _
            Description:="El contenidor de la graella no pot ser Nothing."
    End If

    Set mContainer = ContainerObject
    Set mTheme = modUI.Theme()

    mGridName = Trim$(GridName)

    If Len(mGridName) = 0 Then
        mGridName = "WeeklyGrid"
    End If

    mLeft = LeftValue
    mTop = TopValue

End Sub

'===============================================================================
' Configures the grid dimensions.
'===============================================================================
Public Sub SetDimensions( _
    Optional ByVal NameWidth As Single = DEFAULT_NAME_WIDTH, _
    Optional ByVal StatusWidth As Single = DEFAULT_STATUS_WIDTH, _
    Optional ByVal CellWidth As Single = DEFAULT_CELL_WIDTH, _
    Optional ByVal RowHeight As Single = DEFAULT_ROW_HEIGHT, _
    Optional ByVal HeaderHeight As Single = DEFAULT_HEADER_HEIGHT)

    mNameWidth = NameWidth
    mStatusWidth = StatusWidth
    mCellWidth = CellWidth
    mRowHeight = RowHeight
    mHeaderHeight = HeaderHeight

End Sub

'===============================================================================
' Creates the seven weekday headers from one Monday date.
'===============================================================================
Public Sub SetWeekHeaders(ByVal WeekMonday As Date)

    Dim dayOffset As Long
    Dim currentDate As Date
    Dim captionText As String

    Set mHeaders = New Collection

    For dayOffset = 0 To 6

        currentDate = DateAdd("d", dayOffset, DateValue(WeekMonday))

        captionText = _
            WeekdayShortName(currentDate) & vbCrLf & _
            Format$(currentDate, "dd/mm")

        mHeaders.Add captionText

    Next dayOffset

End Sub

'===============================================================================
' Adds one row.
'
' ShiftCodes must contain one value per visible day.
' Example:
'   Array("M", "M", "M", "I", "I", "I", "I")
'===============================================================================
Public Sub AddRow( _
    ByVal RowName As String, _
    ByVal RowState As String, _
    ByVal ShiftCodes As Variant)

    Dim rowValues As Collection
    Dim index As Long
    Dim firstIndex As Long
    Dim lastIndex As Long

    Set rowValues = New Collection

    If IsArray(ShiftCodes) Then

        firstIndex = LBound(ShiftCodes)
        lastIndex = UBound(ShiftCodes)

        For index = firstIndex To lastIndex
            rowValues.Add CStr(ShiftCodes(index))
        Next index

    Else
        rowValues.Add CStr(ShiftCodes)
    End If

    mRowNames.Add Trim$(RowName)
    mRowStates.Add Trim$(RowState)
    mRows.Add rowValues

End Sub

'===============================================================================
' Removes the stored row data.
'===============================================================================
Public Sub ClearData()

    Set mRowNames = New Collection
    Set mRowStates = New Collection
    Set mRows = New Collection

End Sub

'===============================================================================
' Removes all dynamic controls created by this grid.
'===============================================================================
Public Sub ClearRenderedControls()

    Dim index As Long
    Dim control As Object
    Dim expectedTag As String

    EnsureInitialized

    expectedTag = GridTag()

    For index = mContainer.Controls.Count - 1 To 0 Step -1

        Set control = mContainer.Controls(index)

        If CStr(control.Tag) = expectedTag Then
            mContainer.Controls.Remove control.Name
        End If

    Next index

End Sub

'===============================================================================
' Renders the complete grid.
'===============================================================================
Public Sub Render()

    Dim columnIndex As Long
    Dim rowIndex As Long

    Dim rowValues As Collection

    Dim currentLeft As Single
    Dim currentTop As Single

    EnsureInitialized
    ClearRenderedControls

    RenderHeaderCell _
        ControlName:="NameHeader", _
        CaptionText:="OPERARI", _
        LeftValue:=mLeft, _
        TopValue:=mTop, _
        WidthValue:=mNameWidth

    RenderHeaderCell _
        ControlName:="StateHeader", _
        CaptionText:="ESTAT", _
        LeftValue:=mLeft + mNameWidth + DEFAULT_GAP, _
        TopValue:=mTop, _
        WidthValue:=mStatusWidth

    currentLeft = _
        mLeft + _
        mNameWidth + DEFAULT_GAP + _
        mStatusWidth + DEFAULT_GAP

    For columnIndex = 1 To mHeaders.Count

        RenderHeaderCell _
            ControlName:="DayHeader" & CStr(columnIndex), _
            CaptionText:=CStr(mHeaders(columnIndex)), _
            LeftValue:=currentLeft, _
            TopValue:=mTop, _
            WidthValue:=mCellWidth

        currentLeft = currentLeft + mCellWidth + DEFAULT_GAP

    Next columnIndex

    currentTop = mTop + mHeaderHeight + DEFAULT_GAP

    For rowIndex = 1 To mRows.Count

        RenderNameCell _
            ControlName:="Name" & CStr(rowIndex), _
            CaptionText:=CStr(mRowNames(rowIndex)), _
            LeftValue:=mLeft, _
            TopValue:=currentTop

        RenderStateCell _
            ControlName:="State" & CStr(rowIndex), _
            CaptionText:=CStr(mRowStates(rowIndex)), _
            LeftValue:=mLeft + mNameWidth + DEFAULT_GAP, _
            TopValue:=currentTop

        Set rowValues = mRows(rowIndex)

        currentLeft = _
            mLeft + _
            mNameWidth + DEFAULT_GAP + _
            mStatusWidth + DEFAULT_GAP

        For columnIndex = 1 To mHeaders.Count

            If columnIndex <= rowValues.Count Then

                RenderShiftCell _
                    ControlName:="Shift" & CStr(rowIndex) & "_" & CStr(columnIndex), _
                    ShiftCode:=CStr(rowValues(columnIndex)), _
                    LeftValue:=currentLeft, _
                    TopValue:=currentTop

            Else

                RenderShiftCell _
                    ControlName:="Shift" & CStr(rowIndex) & "_" & CStr(columnIndex), _
                    ShiftCode:=vbNullString, _
                    LeftValue:=currentLeft, _
                    TopValue:=currentTop

            End If

            currentLeft = currentLeft + mCellWidth + DEFAULT_GAP

        Next columnIndex

        currentTop = currentTop + mRowHeight + DEFAULT_GAP

    Next rowIndex

End Sub

'===============================================================================
' Returns the total rendered width.
'===============================================================================
Public Property Get TotalWidth() As Single

    TotalWidth = _
        mNameWidth + DEFAULT_GAP + _
        mStatusWidth + DEFAULT_GAP + _
        (mHeaders.Count * (mCellWidth + DEFAULT_GAP))

End Property

'===============================================================================
' Returns the total rendered height.
'===============================================================================
Public Property Get TotalHeight() As Single

    TotalHeight = _
        mHeaderHeight + DEFAULT_GAP + _
        (mRows.Count * (mRowHeight + DEFAULT_GAP))

End Property

'===============================================================================
' Rendering helpers
'===============================================================================
Private Sub RenderHeaderCell( _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single, _
    ByVal WidthValue As Single)

    Dim label As MSForms.Label

    Set label = CreateLabel(ControlName)

    With label
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = WidthValue
        .Height = mHeaderHeight
        .BackColor = mTheme.PrimaryDarkColor
        .ForeColor = mTheme.SurfaceColor
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
    End With

End Sub

Private Sub RenderNameCell( _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single)

    Dim label As MSForms.Label

    Set label = CreateLabel(ControlName)

    With label
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = mNameWidth
        .Height = mRowHeight
        .BackColor = mTheme.SurfaceColor
        .ForeColor = mTheme.TextColor
        .Font.Bold = True
        .TextAlign = fmTextAlignLeft
        .BorderStyle = fmBorderStyleSingle
    End With

End Sub

Private Sub RenderStateCell( _
    ByVal ControlName As String, _
    ByVal CaptionText As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single)

    Dim label As MSForms.Label

    Set label = CreateLabel(ControlName)

    With label
        .Caption = CaptionText
        .Left = LeftValue
        .Top = TopValue
        .Width = mStatusWidth
        .Height = mRowHeight
        .BackColor = StateBackColor(CaptionText)
        .ForeColor = StateTextColor(CaptionText)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
    End With

End Sub

Private Sub RenderShiftCell( _
    ByVal ControlName As String, _
    ByVal ShiftCode As String, _
    ByVal LeftValue As Single, _
    ByVal TopValue As Single)

    Dim label As MSForms.Label
    Dim normalizedCode As String

    normalizedCode = UCase$(Trim$(ShiftCode))

    Set label = CreateLabel(ControlName)

    With label
        .Caption = normalizedCode
        .Left = LeftValue
        .Top = TopValue
        .Width = mCellWidth
        .Height = mRowHeight
        .BackColor = ShiftBackColor(normalizedCode)
        .ForeColor = ShiftTextColor(normalizedCode)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .ControlTipText = ShiftDescription(normalizedCode)
    End With

End Sub

Private Function CreateLabel(ByVal ControlName As String) As MSForms.Label

    Dim label As MSForms.Label
    Dim completeName As String

    completeName = _
        GRID_CONTROL_PREFIX & _
        mGridName & "_" & _
        ControlName

    Set label = mContainer.Controls.Add( _
        "Forms.Label.1", _
        completeName, _
        True)

    With label
        .Tag = GridTag()
        .Font.Name = mTheme.FontName
        .Font.Size = mTheme.FontSize
        .BackStyle = fmBackStyleOpaque
        .SpecialEffect = fmSpecialEffectFlat
    End With

    Set CreateLabel = label

End Function

'===============================================================================
' Color mapping
'===============================================================================
Private Function ShiftBackColor(ByVal ShiftCode As String) As Long

    Select Case ShiftCode

        Case "M"
            ShiftBackColor = RGB(189, 215, 238)

        Case "T"
            ShiftBackColor = RGB(248, 203, 173)

        Case "I"
            ShiftBackColor = mTheme.PrimaryColor

        Case "D"
            ShiftBackColor = RGB(217, 217, 217)

        Case "B"
            ShiftBackColor = RGB(244, 177, 177)

        Case "V"
            ShiftBackColor = RGB(198, 224, 180)

        Case Else
            ShiftBackColor = mTheme.SurfaceColor

    End Select

End Function

Private Function ShiftTextColor(ByVal ShiftCode As String) As Long

    If ShiftCode = "I" Then
        ShiftTextColor = mTheme.SurfaceColor
    Else
        ShiftTextColor = mTheme.TextColor
    End If

End Function

Private Function StateBackColor(ByVal StateText As String) As Long

    Select Case UCase$(Trim$(StateText))

        Case "ACTIU", "ACTIVA"
            StateBackColor = RGB(226, 239, 218)

        Case "BAIXA"
            StateBackColor = RGB(244, 177, 177)

        Case "VACANCES"
            StateBackColor = RGB(198, 224, 180)

        Case "DESCANS"
            StateBackColor = RGB(217, 217, 217)

        Case Else
            StateBackColor = mTheme.SurfaceColor

    End Select

End Function

Private Function StateTextColor(ByVal StateText As String) As Long

    Select Case UCase$(Trim$(StateText))

        Case "ACTIU", "ACTIVA"
            StateTextColor = mTheme.SuccessColor

        Case "BAIXA"
            StateTextColor = mTheme.ErrorColor

        Case Else
            StateTextColor = mTheme.TextColor

    End Select

End Function

Private Function ShiftDescription(ByVal ShiftCode As String) As String

    Select Case ShiftCode

        Case "M"
            ShiftDescription = "Torn de mati"

        Case "T"
            ShiftDescription = "Torn de tarda"

        Case "I"
            ShiftDescription = "Torn intensiu"

        Case "D"
            ShiftDescription = "Descans"

        Case "B"
            ShiftDescription = "Baixa"

        Case "V"
            ShiftDescription = "Vacances"

        Case Else
            ShiftDescription = "Sense assignacio"

    End Select

End Function

'===============================================================================
' General helpers
'===============================================================================
Private Function WeekdayShortName(ByVal Value As Date) As String

    Select Case Weekday(Value, vbMonday)

        Case 1
            WeekdayShortName = "Dl"

        Case 2
            WeekdayShortName = "Dm"

        Case 3
            WeekdayShortName = "Dc"

        Case 4
            WeekdayShortName = "Dj"

        Case 5
            WeekdayShortName = "Dv"

        Case 6
            WeekdayShortName = "Ds"

        Case 7
            WeekdayShortName = "Dg"

    End Select

End Function

Private Function GridTag() As String

    GridTag = GRID_CONTROL_PREFIX & mGridName

End Function

Private Sub EnsureInitialized()

    If mContainer Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 5001, _
            Source:="clsGrid.EnsureInitialized", _
            Description:="La graella no ha estat inicialitzada."
    End If

End Sub
