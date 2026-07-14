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
Public Sub GenerateSchedule(ByVal context As clsScheduleContext)

    Dim intensiveStartDates As Collection

    On Error GoTo ErrorHandler

    If mIsGenerating Then
        RaiseRotationError _
            ErrorNumber:=ERR_GENERATION_IN_PROGRESS, _
            ProcedureName:="GenerateSchedule", _
            Description:="Ja hi ha una generació de torns en curs."
    End If

    mIsGenerating = True

    ValidateScheduleContext context
    PrepareScheduleGeneration context

    GenerateBaseRotation context

    Set intensiveStartDates = GetIntensiveStartDates(context)

    InsertIntensiveWeeks context, intensiveStartDates

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
' DEBUG
' Shows which employee would be selected for every intensive week.
'===============================================================================
Private Sub InsertIntensiveWeeks( _
    ByVal context As clsScheduleContext, _
    ByVal IntensiveWeeks As Collection)

    Dim intensiveDate As Variant
    Dim employee As clsOperari

    Debug.Print String(70, "-")
    Debug.Print "INTENSIVE ROTATION PREVIEW"
    Debug.Print String(70, "-")

    For Each intensiveDate In IntensiveWeeks

        Set employee = SelectNextIntensiveOperator(context)

        If employee Is Nothing Then

            Debug.Print Format$(intensiveDate, "dd/mm/yyyy"); _
                        " -> CAP OPERARI DISPONIBLE"

        Else

    Debug.Print Format$(intensiveDate, "dd/mm/yyyy"); _
                " -> "; employee.Name

    ' Actualitza l'estat de l'operari perquè
    ' la següent selecció tingui en compte aquest intensiu.
    employee.IntensiveCount = employee.IntensiveCount + 1
    employee.LastIntensiveDate = CDate(intensiveDate)

End If

    Next intensiveDate

End Sub



'===============================================================================
' Validates every dependency required before schedule generation.
'===============================================================================
Public Sub ValidateScheduleContext(ByVal context As clsScheduleContext)

    On Error GoTo ErrorHandler

    If context Is Nothing Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONTEXT_REQUIRED, _
            ProcedureName:="ValidateScheduleContext", _
            Description:="El context de planificacio no pot ser Nothing."
    End If

    ValidateConfiguration context
    ValidatePlanningPeriod context
    ValidateEmployees context

    If Not context.IsReadyForScheduling Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONTEXT_REQUIRED, _
            ProcedureName:="ValidateScheduleContext", _
            Description:="El context no esta preparat per generar la planificacio."
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
Private Sub PrepareScheduleGeneration(ByVal context As clsScheduleContext)

    context.ClearAssignments

End Sub

'===============================================================================
' Generates the base Morning/Afternoon cycle for every active employee.
'
' The cycle length comes from the active configuration.
' Intensive periods are not applied in this version.
'===============================================================================
Private Sub GenerateBaseRotation(ByVal context As clsScheduleContext)

    Dim currentDate As Date
    Dim weekIndex As Long
    Dim cyclePosition As Long
    Dim cycleLength As Long
    Dim ShiftCode As String

    Dim employeeItem As Variant
    Dim employee As clsOperari

    cycleLength = _
        context.configuration.MorningWeeks + _
        context.configuration.AfternoonWeeks

    currentDate = context.StartDate

    Do While currentDate <= context.EndDate

        weekIndex = DateDiff( _
            Interval:="ww", _
            Date1:=context.StartDate, _
            Date2:=currentDate, _
            FirstDayOfWeek:=vbMonday)

        cyclePosition = weekIndex Mod cycleLength

        If cyclePosition < context.configuration.MorningWeeks Then
            ShiftCode = SHIFT_CODE_MORNING
        Else
            ShiftCode = SHIFT_CODE_AFTERNOON
        End If

        For Each employeeItem In context.Employees

            Set employee = employeeItem

            If employee.IsActive Then
                CreateDailyAssignment _
                    context:=context, _
                    employee:=employee, _
                    AssignmentDate:=currentDate, _
                    ShiftCode:=ShiftCode
            End If

        Next employeeItem

        currentDate = DateAdd("d", 1, currentDate)

    Loop

End Sub

'===============================================================================
' Creates one shift assignment in memory.
'===============================================================================
Private Sub CreateDailyAssignment( _
    ByVal context As clsScheduleContext, _
    ByVal employee As clsOperari, _
    ByVal AssignmentDate As Date, _
    ByVal ShiftCode As String)

    Dim assignment As clsShiftAssignment

    Set assignment = New clsShiftAssignment

    Set assignment.employee = employee

    assignment.AssignmentDate = AssignmentDate
    assignment.ShiftCode = ShiftCode
    assignment.IsIntensive = False
    assignment.IsPrimaryDuty = False
    assignment.IsBackupDuty = False

    context.AddAssignment assignment

    Set assignment = Nothing

End Sub

'===============================================================================
' Creates one intensive assignment in memory.
'
' This helper centralizes the creation of intensive assignments so that
' InsertIntensiveWeeks() only coordinates the algorithm.
'===============================================================================
Private Function BuildIntensiveAssignment( _
    ByVal employee As clsOperari, _
    ByVal AssignmentDate As Date) As clsShiftAssignment

    Dim assignment As clsShiftAssignment

    Set assignment = New clsShiftAssignment

    Set assignment.employee = employee

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
    ByVal context As clsScheduleContext) As Collection

    Dim result As Collection
    Dim currentDate As Date

    Set result = New Collection

    currentDate = context.StartDate

    Do While currentDate <= context.EndDate

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
    ByVal context As clsScheduleContext) As clsOperari

    Dim employeeItem As Variant
    Dim employee As clsOperari
    Dim selectedEmployee As clsOperari

    For Each employeeItem In context.Employees

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
Private Sub ValidateConfiguration(ByVal context As clsScheduleContext)

    Dim configuration As clsConfiguracio

    Set configuration = context.configuration

    If configuration Is Nothing Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONFIGURATION_REQUIRED, _
            ProcedureName:="ValidateConfiguration", _
            Description:="El context no conte cap configuracio."
    End If

    If Not configuration.IsValid Then
        RaiseRotationError _
            ErrorNumber:=ERR_CONFIGURATION_INVALID, _
            ProcedureName:="ValidateConfiguration", _
            Description:="La configuracio activa no es valida."
    End If

    Set configuration = Nothing

End Sub

'===============================================================================
' Validates the planning period.
'===============================================================================
Private Sub ValidatePlanningPeriod(ByVal context As clsScheduleContext)

    If Not context.HasPlanningPeriod Then
        RaiseRotationError _
            ErrorNumber:=ERR_PLANNING_PERIOD_REQUIRED, _
            ProcedureName:="ValidatePlanningPeriod", _
            Description:="El periode de planificacio no esta definit."
    End If

    If context.EndDate < context.StartDate Then
        RaiseRotationError _
            ErrorNumber:=ERR_PLANNING_PERIOD_REQUIRED, _
            ProcedureName:="ValidatePlanningPeriod", _
            Description:="El periode de planificacio no es valid."
    End If

End Sub

'===============================================================================
' Validates the employee collection.
'===============================================================================
Private Sub ValidateEmployees(ByVal context As clsScheduleContext)

    Dim employeeItem As Variant
    Dim employee As clsOperari

    If context.EmployeeCount = 0 Then
        RaiseRotationError _
            ErrorNumber:=ERR_EMPLOYEES_REQUIRED, _
            ProcedureName:="ValidateEmployees", _
            Description:="Cal almenys un operari per generar la planificacio."
    End If

    For Each employeeItem In context.Employees

        If Not IsObject(employeeItem) Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="La colleccio conte un element que no es un objecte."
        End If

        If Not TypeOf employeeItem Is clsOperari Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="La colleccio conte un objecte que no es un operari."
        End If

        Set employee = employeeItem

        If Not employee.IsValid Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="Hi ha un operari amb dades obligatories incompletes."
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


