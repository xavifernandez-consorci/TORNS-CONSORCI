VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const CLASS_NAME As String = "clsDashboard"
Private Const ERR_FORM_REQUIRED As Long = vbObjectError + 6000
Private Const ERR_NOT_INITIALIZED As Long = vbObjectError + 6001

Private mForm As Object
Private mContext As clsScheduleContext
Private mIsInitialized As Boolean

Public Sub Initialize(ByVal MainForm As Object)

    If MainForm Is Nothing Then
        Err.Raise _
            Number:=ERR_FORM_REQUIRED, _
            Source:=CLASS_NAME & ".Initialize", _
            Description:="El formulari principal no pot ser Nothing."
    End If

    Set mForm = MainForm
    mIsInitialized = True

End Sub

Public Property Set Context(ByVal Value As clsScheduleContext)

    Set mContext = Value

End Property

Public Property Get Context() As clsScheduleContext

    Set Context = mContext

End Property

Public Property Get HasContext() As Boolean

    HasContext = Not mContext Is Nothing

End Property

Public Sub Refresh()

    EnsureInitialized
    modUIBuilder.BuildMainForm mForm

End Sub

Public Sub Shutdown()

    modUIBuilder.ReleaseMainForm

    Set mContext = Nothing
    Set mForm = Nothing
    mIsInitialized = False

End Sub

Private Sub EnsureInitialized()

    If Not mIsInitialized Or mForm Is Nothing Then
        Err.Raise _
            Number:=ERR_NOT_INITIALIZED, _
            Source:=CLASS_NAME & ".EnsureInitialized", _
            Description:="El dashboard encara no ha estat inicialitzat."
    End If

End Sub
