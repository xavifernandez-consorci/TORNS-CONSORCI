Attribute VB_Name = "modConfiguration"

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modConfiguration
' Layer:        Application / Configuration
'
' Purpose:
'   Manage the active application configuration.
'
' Responsibilities:
'   - Load configuration during application startup.
'   - Expose the active configuration to other modules.
'   - Save configuration through the data-access layer.
'   - Release configuration during application shutdown.
'
' Restrictions:
'   - No direct worksheet or range access.
'   - No scheduling business logic.
'   - No hardcoded business values.
'===============================================================================

Private Const MODULE_NAME As String = "modConfiguration"

Private mConfiguration As clsConfiguracio
Private mIsInitialized As Boolean

Public Sub InitializeConfiguration()
    On Error GoTo ErrorHandler

    If mIsInitialized Then Exit Sub

    Set mConfiguration = modDataAccess.LoadConfiguration()

    If mConfiguration Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 1000, _
            Source:=MODULE_NAME & ".InitializeConfiguration", _
            Description:="No s'ha pogut carregar la configuració de l'aplicació."
    End If

    mIsInitialized = True
    Exit Sub

ErrorHandler:
    Set mConfiguration = Nothing
    mIsInitialized = False

    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".InitializeConfiguration", _
        Description:=Err.Description
End Sub

Public Sub SaveConfiguration()
    On Error GoTo ErrorHandler

    EnsureConfigurationIsInitialized

    modDataAccess.SaveConfiguration mConfiguration
    Exit Sub

ErrorHandler:
    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".SaveConfiguration", _
        Description:=Err.Description
End Sub

Public Sub ReplaceConfiguration(ByVal NewConfiguration As clsConfiguracio)
    On Error GoTo ErrorHandler

    If NewConfiguration Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 1001, _
            Source:=MODULE_NAME & ".ReplaceConfiguration", _
            Description:="La nova configuració no pot ser Nothing."
    End If

    Set mConfiguration = NewConfiguration
    mIsInitialized = True
    Exit Sub

ErrorHandler:
    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".ReplaceConfiguration", _
        Description:=Err.Description
End Sub

Public Sub ReleaseConfiguration()
    Set mConfiguration = Nothing
    mIsInitialized = False
End Sub

Public Property Get CurrentConfiguration() As clsConfiguracio
    EnsureConfigurationIsInitialized

    Set CurrentConfiguration = mConfiguration
End Property

Public Property Get IsConfigurationInitialized() As Boolean
    IsConfigurationInitialized = mIsInitialized
End Property

Private Sub EnsureConfigurationIsInitialized()
    If Not mIsInitialized Or mConfiguration Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 1002, _
            Source:=MODULE_NAME & ".EnsureConfigurationIsInitialized", _
            Description:="La configuració de l'aplicació no està inicialitzada."
    End If
End Sub