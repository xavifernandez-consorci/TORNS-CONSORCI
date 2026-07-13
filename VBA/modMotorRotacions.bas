Attribute VB_Name = "modMotorRotacions"

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    modMotorRotacions
' Layer:        Service
'
' Purpose:
'   Coordinate generation of employee shift rotations.
'
' Current phase:
'   Phase 1 - Validate and prepare the scheduling context.
'
' Responsibilities:
'   - Validate the schedule context.
'   - Prepare an in-memory generation operation.
'   - Clear previously generated in-memory assignments.
'   - Expose the public scheduling entry point.
'
' Restrictions:
'   - No worksheet or range access.
'   - No automatic persistence.
'   - No calendar modification.
'   - Duty assignments belong to modGuardies.
'===============================================================================

Private Const MODULE_NAME As String = "modMotorRotacions"

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
' This first implementation validates and prepares the context only.
' Subsequent phases will append generated assignments to Context.Assignments.
'
' Nothing is written to worksheets or persisted automatically.
'===============================================================================
Public Sub GenerateSchedule(ByVal Context As clsScheduleContext)
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

    ' Phase 2:
    ' GenerateBaseRotation Context
    '
    ' Phase 3:
    ' ApplyIntensiveRotation Context
    '
    ' Phase 4:
    ' ApplyPostIntensiveSequence Context
    '
    ' Duty assignments will be coordinated through modGuardies.

CleanExit:
    mIsGenerating = False
    Exit Sub

ErrorHandler:
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
' Clears prior in-memory results before starting a new generation.
'
' This operation does not modify worksheets or persisted calendar data.
'===============================================================================
Private Sub PrepareScheduleGeneration(ByVal Context As clsScheduleContext)
    Context.ClearAssignments
End Sub

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
End Sub

'===============================================================================
' Validates that the planning period exists and has a valid date range.
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
                Description:="La col·lecció conté un element que no és un operari."
        End If

        If Not TypeOf employeeItem Is clsOperari Then
            RaiseRotationError _
                ErrorNumber:=ERR_INVALID_EMPLOYEE, _
                ProcedureName:="ValidateEmployees", _
                Description:="La col·lecció conté un objecte d'un tipus incorrecte."
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