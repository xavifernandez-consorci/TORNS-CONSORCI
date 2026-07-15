VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    clsAppInfo
' Layer:        Presentation
'
' Version:      1.0
'
' Purpose:
'   Centralize descriptive information about the application.
'
' Responsibilities:
'   - Expose the application name and version.
'   - Expose the current environment.
'   - Expose the expected test count shown by the interface.
'   - Provide formatted labels for UserForms.
'
' Restrictions:
'   - No business logic.
'   - No worksheet access.
'   - No persistence.
'   - No Git or network access.
'===============================================================================

Private Const APPLICATION_NAME As String = "TORNS CONSORCI"
Private Const APPLICATION_DESCRIPTION As String = _
    "Gestio de torns dels operaris del Consorci del Bages"

Private Const VERSION_MAJOR As Long = 2
Private Const VERSION_MINOR As Long = 0
Private Const VERSION_PATCH As Long = 0

Private Const DEFAULT_ENVIRONMENT As String = "Desenvolupament"
Private Const EXPECTED_TEST_COUNT As Long = 45

Private mEnvironmentName As String
Private mBuildLabel As String
Private mLastKnownTestsPassed As Long
Private mLastKnownTestsRun As Long

Private Sub Class_Initialize()

    mEnvironmentName = DEFAULT_ENVIRONMENT
    mBuildLabel = vbNullString
    mLastKnownTestsPassed = EXPECTED_TEST_COUNT
    mLastKnownTestsRun = EXPECTED_TEST_COUNT

End Sub

Public Property Get Name() As String

    Name = APPLICATION_NAME

End Property

Public Property Get Description() As String

    Description = APPLICATION_DESCRIPTION

End Property

Public Property Get VersionMajor() As Long

    VersionMajor = VERSION_MAJOR

End Property

Public Property Get VersionMinor() As Long

    VersionMinor = VERSION_MINOR

End Property

Public Property Get VersionPatch() As Long

    VersionPatch = VERSION_PATCH

End Property

Public Property Get VersionText() As String

    VersionText = _
        CStr(VERSION_MAJOR) & "." & _
        CStr(VERSION_MINOR) & "." & _
        CStr(VERSION_PATCH)

End Property

Public Property Get DisplayTitle() As String

    DisplayTitle = APPLICATION_NAME & " v" & VersionText

End Property

Public Property Get EnvironmentName() As String

    EnvironmentName = mEnvironmentName

End Property

Public Property Let EnvironmentName(ByVal Value As String)

    Dim normalizedValue As String

    normalizedValue = Trim$(Value)

    If Len(normalizedValue) = 0 Then
        mEnvironmentName = DEFAULT_ENVIRONMENT
    Else
        mEnvironmentName = normalizedValue
    End If

End Property

Public Property Get BuildLabel() As String

    BuildLabel = mBuildLabel

End Property

Public Property Let BuildLabel(ByVal Value As String)

    mBuildLabel = Trim$(Value)

End Property

Public Property Get LastKnownTestsPassed() As Long

    LastKnownTestsPassed = mLastKnownTestsPassed

End Property

Public Property Get LastKnownTestsRun() As Long

    LastKnownTestsRun = mLastKnownTestsRun

End Property

Public Property Get TestsStatusText() As String

    TestsStatusText = _
        CStr(mLastKnownTestsPassed) & "/" & _
        CStr(mLastKnownTestsRun) & " tests"

End Property

Public Property Get TestsArePassing() As Boolean

    TestsArePassing = _
        mLastKnownTestsRun > 0 _
        And mLastKnownTestsPassed = mLastKnownTestsRun

End Property

Public Sub UpdateTestStatus( _
    ByVal TestsPassed As Long, _
    ByVal TestsRun As Long)

    If TestsPassed < 0 Then TestsPassed = 0
    If TestsRun < 0 Then TestsRun = 0

    If TestsPassed > TestsRun Then
        TestsPassed = TestsRun
    End If

    mLastKnownTestsPassed = TestsPassed
    mLastKnownTestsRun = TestsRun

End Sub

Public Function Clone() As clsAppInfo

    Dim infoCopy As clsAppInfo

    Set infoCopy = New clsAppInfo

    infoCopy.EnvironmentName = mEnvironmentName
    infoCopy.BuildLabel = mBuildLabel

    infoCopy.UpdateTestStatus _
        TestsPassed:=mLastKnownTestsPassed, _
        TestsRun:=mLastKnownTestsRun

    Set Clone = infoCopy

End Function
