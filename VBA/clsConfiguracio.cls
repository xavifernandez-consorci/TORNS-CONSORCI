Option Explicit

' Purpose: Configuration domain object.
' Responsibilities:
' - Encapsulate configurable application settings.
' - Expose typed accessors for the rest of the application.
' - Remain independent from worksheet storage details.

Public EmployeeCount As Long
Public MorningWeeks As Long
Public AfternoonWeeks As Long
Public IntensiveStartDay As String
Public IntensiveEndDay As String
Public WorkingHoursPerShift As Double
Public MorningStartTime As String
Public MorningEndTime As String
Public AfternoonStartTime As String
Public AfternoonEndTime As String
Public DutyBackupCount As Long
