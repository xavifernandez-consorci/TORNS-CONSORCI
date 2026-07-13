Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modDataAccess
' Layer:        Data
'
' Purpose:
'   Centralize persistence operations between Excel storage and domain objects.
'
' Responsibilities:
'   - Load and save configuration.
'   - Load employees.
'   - Load assignments.
'   - Save assignments.
'
' Restrictions:
'   - No scheduling logic.
'   - No rotation logic.
'   - No guard-duty logic.
'===============================================================================

Public Function LoadConfiguration() As clsConfiguracio

    ' TODO:
    ' Read configuration from workbook storage.
    ' Temporary implementation until persistence is developed.

    Set LoadConfiguration = New clsConfiguracio

End Function

Public Sub SaveConfiguration(ByVal Configuration As clsConfiguracio)

    If Configuration Is Nothing Then Exit Sub

    ' TODO:
    ' Persist configuration to workbook storage.

End Sub

Public Sub LoadEmployees()

    ' TODO:
    ' Read employees from workbook storage.

End Sub

Public Sub LoadAssignments()

    ' TODO:
    ' Read generated schedule from workbook storage.

End Sub

Public Sub SaveAssignments()

    ' TODO:
    ' Persist generated schedule.

End Sub