VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    ThisWorkbook
' Layer:        Workbook
'
' Purpose:
'   Manage the Excel workbook lifecycle and delegate application startup
'   and shutdown operations to modApplicationBootstrap.
'
' Notes:
'   - This component must not contain business logic.
'   - This component must not access worksheets directly.
'   - Errors are delegated to the bootstrap layer whenever possible.
'===============================================================================

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    modApplicationBootstrap.InitializeApplication
    Exit Sub

ErrorHandler:
    modApplicationBootstrap.HandleStartupError _
        Source:="ThisWorkbook.Workbook_Open", _
        ErrorNumber:=Err.Number, _
        ErrorDescription:=Err.Description
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler

    modApplicationBootstrap.ShutdownApplication
    Exit Sub

ErrorHandler:
    modApplicationBootstrap.HandleShutdownError _
        Source:="ThisWorkbook.Workbook_BeforeClose", _
        ErrorNumber:=Err.Number, _
        ErrorDescription:=Err.Description
End Sub