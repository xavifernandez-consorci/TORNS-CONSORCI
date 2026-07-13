Attribute VB_Name = "modApplicationBootstrap"

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modApplicationBootstrap
' Layer:        Application
'
' Purpose:
'   Coordinate application startup and shutdown.
'
' Responsibilities:
'   - Initialize the application state.
'   - Load configuration through the configuration layer.
'   - Open the main form.
'   - Centralize startup and shutdown error reporting.
'
' Restrictions:
'   - No scheduling business logic.
'   - No direct worksheet access.
'   - No hardcoded business values.
'===============================================================================

Private Const MODULE_NAME As String = "modApplicationBootstrap"

Private mIsInitialized As Boolean
Private mIsShuttingDown As Boolean

Public Sub InitializeApplication()
    On Error GoTo ErrorHandler

    If mIsInitialized Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Inicialitzant TORNS CONSORCI..."

    modConfiguration.InitializeConfiguration

    mIsInitialized = True

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    frmInici.Show
    Exit Sub

ErrorHandler:
    RestoreApplicationState

    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".InitializeApplication", _
        Description:=Err.Description
End Sub

Public Sub ShutdownApplication()
    On Error GoTo ErrorHandler

    If mIsShuttingDown Then Exit Sub

    mIsShuttingDown = True

    modConfiguration.ReleaseConfiguration

    RestoreApplicationState

    mIsInitialized = False
    mIsShuttingDown = False
    Exit Sub

ErrorHandler:
    RestoreApplicationState

    mIsInitialized = False
    mIsShuttingDown = False

    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".ShutdownApplication", _
        Description:=Err.Description
End Sub

Public Sub HandleStartupError( _
    ByVal Source As String, _
    ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String)

    RestoreApplicationState

    MsgBox _
        Prompt:="No s'ha pogut iniciar TORNS CONSORCI." & vbCrLf & vbCrLf & _
                "Origen: " & Source & vbCrLf & _
                "Error " & CStr(ErrorNumber) & ": " & ErrorDescription, _
        Buttons:=vbCritical + vbOKOnly, _
        Title:="Error d'inicialització"
End Sub

Public Sub HandleShutdownError( _
    ByVal Source As String, _
    ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String)

    RestoreApplicationState

    MsgBox _
        Prompt:="S'ha produït un error en tancar TORNS CONSORCI." & _
                vbCrLf & vbCrLf & _
                "Origen: " & Source & vbCrLf & _
                "Error " & CStr(ErrorNumber) & ": " & ErrorDescription, _
        Buttons:=vbExclamation + vbOKOnly, _
        Title:="Error de tancament"
End Sub

Public Property Get IsApplicationInitialized() As Boolean
    IsApplicationInitialized = mIsInitialized
End Property

Private Sub RestoreApplicationState()
    On Error Resume Next

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    On Error GoTo 0
End Sub