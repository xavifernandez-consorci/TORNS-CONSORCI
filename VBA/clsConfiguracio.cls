VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfiguracio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    clsConfiguracio
' Layer:        Domain
'
' Purpose:
'   Represent the configurable settings used by the application.
'
' Responsibilities:
'   - Encapsulate configuration values.
'   - Expose strongly typed properties.
'   - Validate values when they are assigned.
'   - Remain independent from worksheet storage details.
'
' Restrictions:
'   - No worksheet or range access.
'   - No configuration persistence.
'   - No scheduling or duty-generation logic.
'===============================================================================

Private Const CLASS_NAME As String = "clsConfiguracio"

Private Const ERR_INVALID_EMPLOYEE_COUNT As Long = vbObjectError + 1100
Private Const ERR_INVALID_WEEK_COUNT As Long = vbObjectError + 1101
Private Const ERR_INVALID_DAY_NAME As Long = vbObjectError + 1102
Private Const ERR_INVALID_WORKING_HOURS As Long = vbObjectError + 1103
Private Const ERR_INVALID_TIME_VALUE As Long = vbObjectError + 1104
Private Const ERR_INVALID_BACKUP_COUNT As Long = vbObjectError + 1105

Private mEmployeeCount As Long
Private mMorningWeeks As Long
Private mAfternoonWeeks As Long
Private mIntensiveStartDay As String
Private mIntensiveEndDay As String
Private mWorkingHoursPerShift As Double
Private mMorningStartTime As String
Private mMorningEndTime As String
Private mAfternoonStartTime As String
Private mAfternoonEndTime As String
Private mDutyBackupCount As Long

Public Property Get EmployeeCount() As Long
    EmployeeCount = mEmployeeCount
End Property

Public Property Let EmployeeCount(ByVal Value As Long)
    If Value <= 0 Then
        RaiseValidationError _
            ErrorNumber:=ERR_INVALID_EMPLOYEE_COUNT, _
            ProcedureName:="EmployeeCount", _
            Description:="El nombre d'operaris ha de ser superior a zero."
    End If

    mEmployeeCount = Value
End Property

Public Property Get MorningWeeks() As Long
    MorningWeeks = mMorningWeeks
End Property

Public Property Let MorningWeeks(ByVal Value As Long)
    ValidateWeekCount Value, "MorningWeeks"
    mMorningWeeks = Value
End Property

Public Property Get AfternoonWeeks() As Long
    AfternoonWeeks = mAfternoonWeeks
End Property

Public Property Let AfternoonWeeks(ByVal Value As Long)
    ValidateWeekCount Value, "AfternoonWeeks"
    mAfternoonWeeks = Value
End Property

Public Property Get IntensiveStartDay() As String
    IntensiveStartDay = mIntensiveStartDay
End Property

Public Property Let IntensiveStartDay(ByVal Value As String)
    mIntensiveStartDay = ValidateRequiredText( _
        Value:=Value, _
        ProcedureName:="IntensiveStartDay", _
        ErrorNumber:=ERR_INVALID_DAY_NAME, _
        FieldDescription:="El dia d'inici de l'intensiu")
End Property

Public Property Get IntensiveEndDay() As String
    IntensiveEndDay = mIntensiveEndDay
End Property

Public Property Let IntensiveEndDay(ByVal Value As String)
    mIntensiveEndDay = ValidateRequiredText( _
        Value:=Value, _
        ProcedureName:="IntensiveEndDay", _
        ErrorNumber:=ERR_INVALID_DAY_NAME, _
        FieldDescription:="El dia de finalització de l'intensiu")
End Property

Public Property Get WorkingHoursPerShift() As Double
    WorkingHoursPerShift = mWorkingHoursPerShift
End Property

Public Property Let WorkingHoursPerShift(ByVal Value As Double)
    If Value <= 0# Or Value > 24# Then
        RaiseValidationError _
            ErrorNumber:=ERR_INVALID_WORKING_HOURS, _
            ProcedureName:="WorkingHoursPerShift", _
            Description:="Les hores de treball per torn han d'estar entre 0 i 24."
    End If

    mWorkingHoursPerShift = Value
End Property

Public Property Get MorningStartTime() As String
    MorningStartTime = mMorningStartTime
End Property

Public Property Let MorningStartTime(ByVal Value As String)
    mMorningStartTime = ValidateTimeText(Value, "MorningStartTime")
End Property

Public Property Get MorningEndTime() As String
    MorningEndTime = mMorningEndTime
End Property

Public Property Let MorningEndTime(ByVal Value As String)
    mMorningEndTime = ValidateTimeText(Value, "MorningEndTime")
End Property

Public Property Get AfternoonStartTime() As String
    AfternoonStartTime = mAfternoonStartTime
End Property

Public Property Let AfternoonStartTime(ByVal Value As String)
    mAfternoonStartTime = ValidateTimeText(Value, "AfternoonStartTime")
End Property

Public Property Get AfternoonEndTime() As String
    AfternoonEndTime = mAfternoonEndTime
End Property

Public Property Let AfternoonEndTime(ByVal Value As String)
    mAfternoonEndTime = ValidateTimeText(Value, "AfternoonEndTime")
End Property

Public Property Get DutyBackupCount() As Long
    DutyBackupCount = mDutyBackupCount
End Property

Public Property Let DutyBackupCount(ByVal Value As Long)
    If Value < 0 Then
        RaiseValidationError _
            ErrorNumber:=ERR_INVALID_BACKUP_COUNT, _
            ProcedureName:="DutyBackupCount", _
            Description:="El nombre de reforços de guàrdia no pot ser negatiu."
    End If

    mDutyBackupCount = Value
End Property

Public Function IsValid() As Boolean
    IsValid = _
        mEmployeeCount > 0 And _
        mMorningWeeks > 0 And _
        mAfternoonWeeks > 0 And _
        Len(mIntensiveStartDay) > 0 And _
        Len(mIntensiveEndDay) > 0 And _
        mWorkingHoursPerShift > 0# And _
        IsDate(mMorningStartTime) And _
        IsDate(mMorningEndTime) And _
        IsDate(mAfternoonStartTime) And _
        IsDate(mAfternoonEndTime) And _
        mDutyBackupCount >= 0
End Function

Private Sub ValidateWeekCount( _
    ByVal Value As Long, _
    ByVal ProcedureName As String)

    If Value <= 0 Then
        RaiseValidationError _
            ErrorNumber:=ERR_INVALID_WEEK_COUNT, _
            ProcedureName:=ProcedureName, _
            Description:="El nombre de setmanes ha de ser superior a zero."
    End If
End Sub

Private Function ValidateRequiredText( _
    ByVal Value As String, _
    ByVal ProcedureName As String, _
    ByVal ErrorNumber As Long, _
    ByVal FieldDescription As String) As String

    Dim normalizedValue As String

    normalizedValue = Trim$(Value)

    If Len(normalizedValue) = 0 Then
        RaiseValidationError _
            ErrorNumber:=ErrorNumber, _
            ProcedureName:=ProcedureName, _
            Description:=FieldDescription & " és obligatori."
    End If

    ValidateRequiredText = normalizedValue
End Function

Private Function ValidateTimeText( _
    ByVal Value As String, _
    ByVal ProcedureName As String) As String

    Dim normalizedValue As String

    normalizedValue = Trim$(Value)

    If Len(normalizedValue) = 0 Or Not IsDate(normalizedValue) Then
        RaiseValidationError _
            ErrorNumber:=ERR_INVALID_TIME_VALUE, _
            ProcedureName:=ProcedureName, _
            Description:="L'horari indicat no és una hora vàlida."
    End If

    ValidateTimeText = Format$(TimeValue(normalizedValue), "hh:mm")
End Function

Private Sub RaiseValidationError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=CLASS_NAME & "." & ProcedureName, _
        Description:=Description
End Sub