Attribute VB_Name = "modMotorRotacions"

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modMotorRotacions
' Layer:        Service
'
' Version:      1.0
'
' Purpose:
'   Coordinate the in-memory generation of employee shift rotations.
'
' Included in this version:
'   - Scheduling context validation.
'   - Safe generation lifecycle.
'   - Base Morning/Afternoon rotation.
'   - Detection of intensive-period start dates.
'   - Selection of the next intensive operator.
'   - Creation of in-memory shift assignments.
'
' Not included yet:
'   - Intensive assignment insertion.
'   - Post-intensive sequence.
'   - Duty backup assignment.
'   - Persistence.
'   - Worksheet or calendar modification.
'
' Restrictions:
'   - No worksheet or range access.
'   - No automatic persistence.
'   - No form or message-box interaction.
'   - Duty backup logic belongs to modGuardies.
'===============================================================================

Private Const MODULE_NAME As String = "modMotorRotacions"

Private Const SHIFT_CODE_MORNING As String = "M"
Private Const SHIFT_CODE_AFTERNOON As String = "T"

Private Const ERR_CONTEXT_REQUIRED As Long = vbObjectError + 2000
Private Const ERR_CONFIGURATION_REQUIRED As Long = vbObjectError + 2001
Private Const ERR_CONFIGURATION_INVALID As Long = vbObjectError + 2002
Private Const ERR_EMPLOYEES_REQUIRED As Long = vbObjectError + 2003
Private Const ERR_INVALID_EMPLOYEE As Long = vbObjectError + 2004
Private Const ERR_PLANNING_PERIOD_REQUIRED As Long = vbObjectError + 2005
Private Const ERR_GENERATION_IN_PROGRESS As Long = vbObjectError + 2006

Private mIsGenerating As Boolean

'===============================================================================
' Public scheduling entry point.
'
' The generated assignments remain in Context.Assignments.
' Nothing is persisted or written to Excel automatically.
'===============================================================================
Public Sub GenerateSchedule(ByVal Context As clsScheduleContext)

    Dim intensiveStartDates As Collection

    On Error GoTo ErrorHandler

    If mIsGenerating Then
        RaiseRotationError _
            ErrorNumber:=ERR_GENERATION_IN_PROGRESS, _
            ProcedureName:="GenerateSchedule", _
            Description:="Ja hi ha una generació de torns en curs."
    End If

    mIsGenerating = True

    ValidateScheduleContext Context
    PrepareScheduleGeneration Context

    GenerateBaseRotation Context

    Set intensiveStartDates = GetIntensiveStartDates(Context)

    ' Version 1.1:
    ' InsertIntensiveWeeks Context, intensiveStartDates
    '
    ' Version 1.2:
    ' ApplyPostIntensiveSequence Context
    '
    ' Version 1.3:
    ' ValidateGeneratedSchedule Context

CleanExit:
    Set intensiveStartDates = Nothing
    mIsGenerating = False
    Exit Sub

ErrorHandler:
    Set intensiveStartDates = Nothing
    mIsGenerating = False

    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".GenerateSchedule", _
        Description:=Err.Description

End Sub

'===============================================================================
' Validates every dependency required before schedule generation.
'===============================================================================
Public Sub ValidateScheduleContext(ByVal Context As clsScheduleContext)

    On Error GoTo ErrorHandler

    If Context Is Nothing Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONTEXT_REQUIRED, _
            ProcedureName:="ValidateScheduleContext", _
            Description:="El context de planificació no pot ser Nothing."
    End If

    ValidateConfiguration Context
    ValidatePlanningPeriod Context
    ValidateEmployees Context

    If Not Context.IsReadyForScheduling Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONTEXT_REQUIRED, _
            ProcedureName:="ValidateScheduleContext", _
            Description:="El context no està preparat per generar la planificació."
    End If

    Exit Sub

ErrorHandler:
    Err.Raise _
        Number:=Err.Number, _
        Source:=MODULE_NAME & ".ValidateScheduleContext", _
        Description:=Err.Description

End Sub

'===============================================================================
' Indicates whether the rotation motor is currently running.
'===============================================================================
Public Property Get IsGenerating() As Boolean

    IsGenerating = mIsGenerating

End Property

'===============================================================================
' Clears previous in-memory results before a new generation.
'
' Persistent calendar data is not modified.
'===============================================================================
Private Sub PrepareScheduleGeneration(ByVal Context As clsScheduleContext)

    Context.ClearAssignments

End Sub

'===============================================================================
' Generates the base Morning/Afternoon cycle for every active employee.
'
' The cycle length comes from the active configuration.
' Intensive periods are not applied in this version.
'===============================================================================
Private Sub GenerateBaseRotation(ByVal Context As clsScheduleContext)

    Dim currentDate As Date
    Dim weekIndex As Long
    Dim cyclePosition As Long
    Dim cycleLength As Long
    Dim shiftCode As String

    Dim employeeItem As Variant
    Dim employee As clsOperari

    cycleLength = _
        Context.Configuration.MorningWeeks + _
        Context.Configuration.AfternoonWeeks

    currentDate = Context.StartDate

    Do While currentDate <= Context.EndDate

        weekIndex = DateDiff( _
            Interval:="ww", _
            Date1:=Context.StartDate, _
            Date2:=currentDate, _
            FirstDayOfWeek:=vbMonday)

        cyclePosition = weekIndex Mod cycleLength

        If cyclePosition < Context.Configuration.MorningWeeks Then
            shiftCode = SHIFT_CODE_MORNING
        Else
            shiftCode = SHIFT_CODE_AFTERNOON
        End If

        For Each employeeItem In Context.Employees

            Set employee = employeeItem

            If employee.IsActive Then
                CreateDailyAssignment _
                    Context:=Context, _
                    Employee:=employee, _
                    AssignmentDate:=currentDate, _
                    ShiftCode:=shiftCode
            End If

        Next employeeItem

        currentDate = DateAdd("d", 1, currentDate)

    Loop

End Sub

'===============================================================================
' Creates one shift assignment in memory.
'===============================================================================
Private Sub CreateDailyAssignment( _
    ByVal Context As clsScheduleContext, _
    ByVal Employee As clsOperari, _
    ByVal AssignmentDate As Date, _
    ByVal ShiftCode As String)

    Dim assignment As clsShiftAssignment

    Set assignment = New clsShiftAssignment

    Set assignment.Employee = Employee

    assignment.AssignmentDate = AssignmentDate
    assignment.ShiftCode = ShiftCode
    assignment.IsIntensive = False
    assignment.IsPrimaryDuty = False
    assignment.IsBackupDuty = False

    Context.AddAssignment assignment

    Set assignment = Nothing

End Sub

'===============================================================================
' Creates one intensive assignment in memory.
'
' This helper centralizes the creation of intensive assignments so that
' InsertIntensiveWeeks() only coordinates the algorithm.
'===============================================================================
Private Function BuildIntensiveAssignment( _
    ByVal Employee As clsOperari, _
    ByVal AssignmentDate As Date) As clsShiftAssignment

    Dim assignment As clsShiftAssignment

    Set assignment = New clsShiftAssignment

    Set assignment.Employee = Employee

    assignment.AssignmentDate = AssignmentDate

    ' TODO:
    ' The intensive shift code will eventually come from clsConfiguracio.
    assignment.ShiftCode = "I"

    assignment.IsIntensive = True
    assignment.IsPrimaryDuty = True
    assignment.IsBackupDuty = False

    Set BuildIntensiveAssignment = assignment

End Function

'===============================================================================
' Returns every Thursday contained in the planning period.
'
' These dates are candidates for the start of an intensive block.
'===============================================================================
Private Function GetIntensiveStartDates( _
    ByVal Context As clsScheduleContext) As Collection

    Dim result As Collection
    Dim currentDate As Date

    Set result = New Collection

    currentDate = Context.StartDate

    Do While currentDate <= Context.EndDate

        If Weekday(currentDate, vbMonday) = 4 Then
            result.Add currentDate
        End If

        currentDate = DateAdd("d", 1, currentDate)

    Loop

    Set GetIntensiveStartDates = result

End Function

'===============================================================================
' Returns the next eligible employee for an intensive rotation.
'
' Selection priority:
'   1. Active employee.
'   2. Intensive candidate.
'   3. Lowest accumulated intensive count.
'   4. Oldest last-intensive date.
'   5. Stable collection order.
'===============================================================================
Private Function SelectNextIntensiveOperator( _
    ByVal Context As clsScheduleContext) As clsOperari

    Dim employeeItem As Variant
    Dim employee As clsOperari
    Dim selectedEmployee As clsOperari

    For Each employeeItem In Context.Employees

        Set employee = employeeItem

        If employee.IsActive And employee.IsIntensiveCandidate Then

            If selectedEmployee Is Nothing Then

                Set selectedEmployee = employee

            ElseIf employee.IntensiveCount < selectedEmployee.IntensiveCount Then

                Set selectedEmployee = employee

            ElseIf employee.IntensiveCount = selectedEmployee.IntensiveCount Then

                If CompareLastIntensive( _
                    FirstEmployee:=employee, _
                    SecondEmployee:=selectedEmployee) < 0 Then

                    Set selectedEmployee = employee

                End If

            End If

        End If

    Next employeeItem

    Set SelectNextIntensiveOperator = selectedEmployee

End Function

'===============================================================================
' Compares employees by their last intensive date.
'
' Return values:
'   -1: FirstEmployee has priority.
'    0: Both have equal priority.
'    1: SecondEmployee has priority.
'
' An employee without previous intensives has priority over an employee that
' already has a registered intensive date.
'===============================================================================
Private Function CompareLastIntensive( _
    ByVal FirstEmployee As clsOperari, _
    ByVal SecondEmployee As clsOperari) As Long

    If Not FirstEmployee.HasLastIntensiveDate Then

        If Not SecondEmployee.HasLastIntensiveDate Then
            CompareLastIntensive = 0
        Else
            CompareLastIntensive = -1
        End If

        Exit Function

    End If

    If Not SecondEmployee.HasLastIntensiveDate Then
        CompareLastIntensive = 1
        Exit Function
    End If

    If FirstEmployee.LastIntensiveDate < SecondEmployee.LastIntensiveDate Then
        CompareLastIntensive = -1

    ElseIf FirstEmployee.LastIntensiveDate > SecondEmployee.LastIntensiveDate Then
        CompareLastIntensive = 1

    Else
        CompareLastIntensive = 0
    End If

End Function

'===============================================================================
' Validates the active configuration.
'===============================================================================
Private Sub ValidateConfiguration(ByVal Context As clsScheduleContext)

    Dim configuration As clsConfiguracio

    Set configuration = Context.Configuration

    If configuration Is Nothing Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONFIGURATION_REQUIRED, _
            ProcedureName:="ValidateConfiguration", _
            Description:="El context no conté cap configuració."
    End If

    If Not configuration.IsValid Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONFIGURATION_INVALID, _
            ProcedureName:="ValidateConfiguration", _
            Description:="La configuració activa no és vàlida."
    End If

    Set configuration = Nothing

End Sub

'===============================================================================
' Validates the planning period.
'===============================================================================
Private Sub ValidatePlanningPeriod(ByVal Context As clsScheduleContext)

    If Not Context.HasPlanningPeriod Then
        RaiseRotationError _
            ErrorNumber:=ERR_PLANNING_PERIOD_REQUIRED, _
            ProcedureName:="ValidatePlanningPeriod", _
            Description:="El període de planificació no està definit."
    End If

    If Context.EndDate < Context.StartDate Then
        RaiseRotationError _
            ErrorNumber:=ERR_PLANNING_PERIOD_REQUIRED, _
            ProcedureName:="ValidatePlanningPeriod", _
            Description:="El període de planificació no és vàlid."
    End If

End Sub

'===============================================================================
' Validates the employee collection.
'===============================================================================
Private Sub ValidateEmployees(ByVal Context As clsScheduleContext)

    Dim employeeItem As Variant
    Dim employee As clsOperari

    If Context.EmployeeCount = 0 Then
        RaiseRotationError _
            ErrorNumber:=ERR_EMPLOYEES_REQUIRED, _
            ProcedureName:="ValidateEmployees", _
            Description:="Cal almenys un operari per generar la planificació."
    End If

    For Each employeeItem In Context.Employees

        If Not IsObject(employeeItem) Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="La col·lecció conté un element que no és un objecte."
        End If

        If Not TypeOf employeeItem Is clsOperari Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="La col·lecció conté un objecte que no és un operari."
        End If

        Set employee = employeeItem

        If Not employee.IsValid Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="Hi ha un operari amb dades obligatòries incompletes."
        End If

    Next employeeItem

End Sub

'===============================================================================
' Raises a controlled service-layer error.
'===============================================================================
Private Sub RaiseRotationError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=MODULE_NAME & "." & ProcedureName, _
        Description:=Description

End Sub