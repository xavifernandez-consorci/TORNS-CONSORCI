Attribute VB_Name = "modSubstitucions"
Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modSubstitucions
' Layer:        Service
'
' Version:      1.0
'
' Purpose:
'   Apply manual substitutions to an already generated schedule without
'   recalculating the remainder of the planning period.
'
' Business rule implemented:
'   - The originally planned intensive operator is marked as BAIXA from Monday
'     through Sunday of the affected week.
'   - The replacement operator is marked as DESCANS from Monday through
'     Wednesday and INTENSIU from Thursday through Sunday.
'   - The replacement does not alter the replacement operator's assignments
'     from the following Monday onward.
'   - Every changed assignment is marked as a manual override.
'
' Restrictions:
'   - No worksheet or range access.
'   - No persistence.
'   - No automatic recalculation of later weeks.
'   - No user-interface interaction.
'===============================================================================

Private Const MODULE_NAME As String = "modSubstitucions"

Private Const SHIFT_CODE_SICK_LEAVE As String = "B"
Private Const SHIFT_CODE_REST As String = "D"
Private Const SHIFT_CODE_INTENSIVE As String = "I"

Private Const DAYS_BEFORE_INTENSIVE As Long = 3
Private Const INTENSIVE_DAY_COUNT As Long = 4
Private Const AFFECTED_WEEK_DAY_COUNT As Long = 7

Private Const ERR_CONTEXT_REQUIRED As Long = vbObjectError + 3000
Private Const ERR_ORIGINAL_EMPLOYEE_REQUIRED As Long = vbObjectError + 3001
Private Const ERR_REPLACEMENT_EMPLOYEE_REQUIRED As Long = vbObjectError + 3002
Private Const ERR_SAME_EMPLOYEE As Long = vbObjectError + 3003
Private Const ERR_INVALID_INTENSIVE_DATE As Long = vbObjectError + 3004
Private Const ERR_REASON_REQUIRED As Long = vbObjectError + 3005
Private Const ERR_ASSIGNMENT_NOT_FOUND As Long = vbObjectError + 3006
Private Const ERR_REPLACEMENT_NOT_AVAILABLE As Long = vbObjectError + 3007

'===============================================================================
' Replaces one planned intensive block because the original operator is on leave.
'
' IntensiveStartDate must be the Thursday on which the intensive block starts.
'
' Result:
'   Original employee:
'       Monday-Sunday = BAIXA
'
'   Replacement employee:
'       Monday-Wednesday = DESCANS
'       Thursday-Sunday  = INTENSIU
'
' Assignments from the following Monday onward remain untouched.
'===============================================================================
Public Sub ReplaceIntensiveBlock( _
    ByVal Context As clsScheduleContext, _
    ByVal OriginalEmployee As clsOperari, _
    ByVal ReplacementEmployee As clsOperari, _
    ByVal IntensiveStartDate As Date, _
    ByVal Reason As String)

    Dim weekMonday As Date
    Dim normalizedReason As String

    On Error GoTo ErrorHandler

    ValidateReplacementRequest _
        Context:=Context, _
        OriginalEmployee:=OriginalEmployee, _
        ReplacementEmployee:=ReplacementEmployee, _
        IntensiveStartDate:=IntensiveStartDate, _
        Reason:=Reason

    normalizedReason = Trim$(Reason)
    weekMonday = DateAdd("d", -DAYS_BEFORE_INTENSIVE, DateValue(IntensiveStartDate))

    ApplySickLeaveWeek _
        Context:=Context, _
        OriginalEmployee:=OriginalEmployee, _
        WeekMonday:=weekMonday, _
        Reason:=normalizedReason

    ApplyReplacementWeek _
        Context:=Context, _
        OriginalEmployee:=OriginalEmployee, _
        ReplacementEmployee:=ReplacementEmployee, _
        WeekMonday:=weekMonday, _
        IntensiveStartDate:=DateValue(IntensiveStartDate), _
        Reason:=normalizedReason

    RegisterPerformedIntensive _
        ReplacementEmployee:=ReplacementEmployee, _
        IntensiveStartDate:=DateValue(IntensiveStartDate)

    Exit Sub

ErrorHandler:
    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".ReplaceIntensiveBlock", _
        Description:=Err.Description
End Sub

'===============================================================================
' Validates the complete substitution request.
'===============================================================================
Private Sub ValidateReplacementRequest( _
    ByVal Context As clsScheduleContext, _
    ByVal OriginalEmployee As clsOperari, _
    ByVal ReplacementEmployee As clsOperari, _
    ByVal IntensiveStartDate As Date, _
    ByVal Reason As String)

    If Context Is Nothing Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_CONTEXT_REQUIRED, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="El context de planificacio no pot ser Nothing."
    End If

    If OriginalEmployee Is Nothing Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_ORIGINAL_EMPLOYEE_REQUIRED, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="Cal indicar l'operari original de l'intensiu."
    End If

    If ReplacementEmployee Is Nothing Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_REPLACEMENT_EMPLOYEE_REQUIRED, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="Cal indicar l'operari substitut."
    End If

    If StrComp( _
        OriginalEmployee.Id, _
        ReplacementEmployee.Id, _
        vbTextCompare) = 0 Then

        RaiseSubstitutionError _
            ErrorNumber:=ERR_SAME_EMPLOYEE, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="L'operari original i el substitut no poden ser el mateix."
    End If

    If Weekday(DateValue(IntensiveStartDate), vbMonday) <> 4 Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_INVALID_INTENSIVE_DATE, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="La data d'inici de l'intensiu ha de ser un dijous."
    End If

    If Len(Trim$(Reason)) = 0 Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_REASON_REQUIRED, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="Cal indicar el motiu de la substitucio."
    End If

    If Not ReplacementEmployee.IsActive Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_REPLACEMENT_NOT_AVAILABLE, _
            ProcedureName:="ValidateReplacementRequest", _
            Description:="L'operari substitut no esta actiu."
    End If
End Sub

'===============================================================================
' Marks the original employee as on sick leave from Monday through Sunday.
'===============================================================================
Private Sub ApplySickLeaveWeek( _
    ByVal Context As clsScheduleContext, _
    ByVal OriginalEmployee As clsOperari, _
    ByVal WeekMonday As Date, _
    ByVal Reason As String)

    Dim dayOffset As Long
    Dim assignmentDate As Date
    Dim replacementAssignment As clsShiftAssignment

    For dayOffset = 0 To AFFECTED_WEEK_DAY_COUNT - 1

        assignmentDate = DateAdd("d", dayOffset, WeekMonday)

        Set replacementAssignment = BuildManualAssignment( _
            Employee:=OriginalEmployee, _
            AssignmentDate:=assignmentDate, _
            ShiftCode:=SHIFT_CODE_SICK_LEAVE, _
            IsIntensive:=False, _
            IsPrimaryDuty:=False, _
            OriginalEmployeeId:=OriginalEmployee.Id, _
            Reason:=Reason)

        ReplaceExistingAssignment _
            Context:=Context, _
            Assignment:=replacementAssignment

        Set replacementAssignment = Nothing

    Next dayOffset
End Sub

'===============================================================================
' Applies the replacement operator's temporary week.
'
' Monday-Wednesday are rest days.
' Thursday-Sunday are intensive days.
' No assignment after Sunday is changed.
'===============================================================================
Private Sub ApplyReplacementWeek( _
    ByVal Context As clsScheduleContext, _
    ByVal OriginalEmployee As clsOperari, _
    ByVal ReplacementEmployee As clsOperari, _
    ByVal WeekMonday As Date, _
    ByVal IntensiveStartDate As Date, _
    ByVal Reason As String)

    Dim dayOffset As Long
    Dim assignmentDate As Date
    Dim shiftCode As String
    Dim isIntensiveDay As Boolean
    Dim replacementAssignment As clsShiftAssignment

    For dayOffset = 0 To AFFECTED_WEEK_DAY_COUNT - 1

        assignmentDate = DateAdd("d", dayOffset, WeekMonday)
        isIntensiveDay = assignmentDate >= IntensiveStartDate

        If isIntensiveDay Then
            shiftCode = SHIFT_CODE_INTENSIVE
        Else
            shiftCode = SHIFT_CODE_REST
        End If

        Set replacementAssignment = BuildManualAssignment( _
            Employee:=ReplacementEmployee, _
            AssignmentDate:=assignmentDate, _
            ShiftCode:=shiftCode, _
            IsIntensive:=isIntensiveDay, _
            IsPrimaryDuty:=isIntensiveDay, _
            OriginalEmployeeId:=OriginalEmployee.Id, _
            Reason:=Reason)

        ReplaceExistingAssignment _
            Context:=Context, _
            Assignment:=replacementAssignment

        Set replacementAssignment = Nothing

    Next dayOffset
End Sub

'===============================================================================
' Builds one assignment carrying manual-override metadata.
'===============================================================================
Private Function BuildManualAssignment( _
    ByVal Employee As clsOperari, _
    ByVal AssignmentDate As Date, _
    ByVal ShiftCode As String, _
    ByVal IsIntensive As Boolean, _
    ByVal IsPrimaryDuty As Boolean, _
    ByVal OriginalEmployeeId As String, _
    ByVal Reason As String) As clsShiftAssignment

    Dim assignment As clsShiftAssignment

    Set assignment = New clsShiftAssignment

    Set assignment.Employee = Employee
    assignment.AssignmentDate = AssignmentDate
    assignment.ShiftCode = ShiftCode
    assignment.IsIntensive = IsIntensive
    assignment.IsPrimaryDuty = IsPrimaryDuty
    assignment.IsBackupDuty = False
    assignment.Notes = Reason

    assignment.MarkAsManualOverride _
        OriginalEmployeeId:=OriginalEmployeeId, _
        OverrideReason:=Reason

    Set BuildManualAssignment = assignment
End Function

'===============================================================================
' Replaces an existing assignment without creating duplicates.
'===============================================================================
Private Sub ReplaceExistingAssignment( _
    ByVal Context As clsScheduleContext, _
    ByVal Assignment As clsShiftAssignment)

    Dim existingAssignment As clsShiftAssignment

    Set existingAssignment = Context.FindAssignment( _
        EmployeeId:=Assignment.Employee.Id, _
        AssignmentDate:=Assignment.AssignmentDate)

    If existingAssignment Is Nothing Then
        RaiseSubstitutionError _
            ErrorNumber:=ERR_ASSIGNMENT_NOT_FOUND, _
            ProcedureName:="ReplaceExistingAssignment", _
            Description:="No existeix l'assignacio que s'ha de substituir."
    End If

    Context.ReplaceAssignment Assignment
End Sub

'===============================================================================
' Records that the replacement operator actually performed the intensive block.
'
' This updates history only. It does not recalculate future assignments.
'===============================================================================
Private Sub RegisterPerformedIntensive( _
    ByVal ReplacementEmployee As clsOperari, _
    ByVal IntensiveStartDate As Date)

    ReplacementEmployee.IntensiveCount = _
        ReplacementEmployee.IntensiveCount + 1

    ReplacementEmployee.LastIntensiveDate = IntensiveStartDate
End Sub

'===============================================================================
' Raises a controlled substitution-service error.
'===============================================================================
Private Sub RaiseSubstitutionError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=MODULE_NAME & "." & ProcedureName, _
        Description:=Description
End Sub
