VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShiftAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    clsShiftAssignment
' Layer:        Domain
'
' Version:      1.1
'
' Purpose:
'   Represent one employee assignment for one planning date.
'
' Responsibilities:
'   - Hold the employee assigned to a date.
'   - Hold the assigned shift.
'   - Hold intensive and duty flags.
'   - Record manual substitutions without recalculating the full schedule.
'   - Preserve the originally planned employee when an override is applied.
'
' Restrictions:
'   - No worksheet access.
'   - No persistence logic.
'   - No rotation-generation logic.
'   - No duty-selection logic.
'===============================================================================

Private Const CLASS_NAME As String = "clsShiftAssignment"

Private Const ERR_EMPLOYEE_REQUIRED As Long = vbObjectError + 1300
Private Const ERR_INVALID_DATE As Long = vbObjectError + 1301
Private Const ERR_SHIFT_REQUIRED As Long = vbObjectError + 1302
Private Const ERR_ORIGINAL_EMPLOYEE_REQUIRED As Long = vbObjectError + 1303
Private Const ERR_OVERRIDE_REASON_REQUIRED As Long = vbObjectError + 1304

Private mEmployee As clsOperari
Private mAssignmentDate As Date
Private mHasAssignmentDate As Boolean
Private mShiftCode As String
Private mIsIntensive As Boolean
Private mIsPrimaryDuty As Boolean
Private mIsBackupDuty As Boolean
Private mNotes As String

Private mIsManualOverride As Boolean
Private mOriginalEmployeeId As String
Private mOverrideReason As String

Public Property Get Employee() As clsOperari
    Set Employee = mEmployee
End Property

Public Property Set Employee(ByVal Value As clsOperari)
    If Value Is Nothing Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_EMPLOYEE_REQUIRED, _
            ProcedureName:="Employee", _
            Description:="L'operari de l'assignacio no pot ser Nothing."
    End If

    Set mEmployee = Value
End Property

Public Property Get AssignmentDate() As Date
    EnsureAssignmentDateIsDefined
    AssignmentDate = mAssignmentDate
End Property

Public Property Let AssignmentDate(ByVal Value As Date)
    If Value <= 0 Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_INVALID_DATE, _
            ProcedureName:="AssignmentDate", _
            Description:="La data de l'assignacio no es valida."
    End If

    mAssignmentDate = DateValue(Value)
    mHasAssignmentDate = True
End Property

Public Property Get HasAssignmentDate() As Boolean
    HasAssignmentDate = mHasAssignmentDate
End Property

Public Property Get ShiftCode() As String
    ShiftCode = mShiftCode
End Property

Public Property Let ShiftCode(ByVal Value As String)
    Dim normalizedValue As String

    normalizedValue = Trim$(Value)

    If Len(normalizedValue) = 0 Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_SHIFT_REQUIRED, _
            ProcedureName:="ShiftCode", _
            Description:="El codi de torn es obligatori."
    End If

    mShiftCode = normalizedValue
End Property

Public Property Get IsIntensive() As Boolean
    IsIntensive = mIsIntensive
End Property

Public Property Let IsIntensive(ByVal Value As Boolean)
    mIsIntensive = Value
End Property

Public Property Get IsPrimaryDuty() As Boolean
    IsPrimaryDuty = mIsPrimaryDuty
End Property

Public Property Let IsPrimaryDuty(ByVal Value As Boolean)
    mIsPrimaryDuty = Value
End Property

Public Property Get IsBackupDuty() As Boolean
    IsBackupDuty = mIsBackupDuty
End Property

Public Property Let IsBackupDuty(ByVal Value As Boolean)
    mIsBackupDuty = Value
End Property

Public Property Get Notes() As String
    Notes = mNotes
End Property

Public Property Let Notes(ByVal Value As String)
    mNotes = Trim$(Value)
End Property

Public Property Get IsManualOverride() As Boolean
    IsManualOverride = mIsManualOverride
End Property

Public Property Get OriginalEmployeeId() As String
    OriginalEmployeeId = mOriginalEmployeeId
End Property

Public Property Get OverrideReason() As String
    OverrideReason = mOverrideReason
End Property

'===============================================================================
' Marks the assignment as a manual substitution.
'
' The Employee property must already contain the replacement employee.
' OriginalEmployeeId identifies the employee initially planned by the motor.
'===============================================================================
Public Sub MarkAsManualOverride( _
    ByVal OriginalEmployeeId As String, _
    ByVal OverrideReason As String)

    Dim normalizedEmployeeId As String
    Dim normalizedReason As String

    normalizedEmployeeId = Trim$(OriginalEmployeeId)
    normalizedReason = Trim$(OverrideReason)

    If Len(normalizedEmployeeId) = 0 Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_ORIGINAL_EMPLOYEE_REQUIRED, _
            ProcedureName:="MarkAsManualOverride", _
            Description:="Cal indicar l'operari planificat original."
    End If

    If Len(normalizedReason) = 0 Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_OVERRIDE_REASON_REQUIRED, _
            ProcedureName:="MarkAsManualOverride", _
            Description:="Cal indicar el motiu del canvi manual."
    End If

    mOriginalEmployeeId = normalizedEmployeeId
    mOverrideReason = normalizedReason
    mIsManualOverride = True
End Sub

'===============================================================================
' Removes manual-override metadata from the assignment.
'===============================================================================
Public Sub ClearManualOverride()
    mIsManualOverride = False
    mOriginalEmployeeId = vbNullString
    mOverrideReason = vbNullString
End Sub

Public Function IsValid() As Boolean
    IsValid = _
        Not mEmployee Is Nothing _
        And mHasAssignmentDate _
        And Len(mShiftCode) > 0

    If IsValid And mIsManualOverride Then
        IsValid = _
            Len(mOriginalEmployeeId) > 0 _
            And Len(mOverrideReason) > 0
    End If
End Function

Public Function Clone() As clsShiftAssignment
    Dim assignmentCopy As clsShiftAssignment

    Set assignmentCopy = New clsShiftAssignment

    If Not mEmployee Is Nothing Then
        Set assignmentCopy.Employee = mEmployee
    End If

    If mHasAssignmentDate Then
        assignmentCopy.AssignmentDate = mAssignmentDate
    End If

    If Len(mShiftCode) > 0 Then
        assignmentCopy.ShiftCode = mShiftCode
    End If

    assignmentCopy.IsIntensive = mIsIntensive
    assignmentCopy.IsPrimaryDuty = mIsPrimaryDuty
    assignmentCopy.IsBackupDuty = mIsBackupDuty
    assignmentCopy.Notes = mNotes

    If mIsManualOverride Then
        assignmentCopy.MarkAsManualOverride _
            OriginalEmployeeId:=mOriginalEmployeeId, _
            OverrideReason:=mOverrideReason
    End If

    Set Clone = assignmentCopy
End Function

Private Sub EnsureAssignmentDateIsDefined()
    If Not mHasAssignmentDate Then
        RaiseAssignmentError _
            ErrorNumber:=ERR_INVALID_DATE, _
            ProcedureName:="EnsureAssignmentDateIsDefined", _
            Description:="La data de l'assignacio encara no esta definida."
    End If
End Sub

Private Sub RaiseAssignmentError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=CLASS_NAME & "." & ProcedureName, _
        Description:=Description
End Sub
