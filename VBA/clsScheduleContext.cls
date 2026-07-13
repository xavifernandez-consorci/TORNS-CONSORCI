VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScheduleContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' Project:      TORNS CONSORCI
' Component:    clsScheduleContext
' Layer:        Domain
'
' Purpose:
'   Encapsulate all data required during schedule generation.
'
' Responsibilities:
'   - Hold the active configuration.
'   - Hold the employee collection.
'   - Define the planning period.
'   - Hold the generated shift assignments.
'
' Restrictions:
'   - No worksheet access.
'   - No persistence logic.
'   - No rotation-generation logic.
'   - No duty-generation logic.
'===============================================================================

Private Const CLASS_NAME As String = "clsScheduleContext"

Private Const ERR_CONFIGURATION_REQUIRED As Long = vbObjectError + 1200
Private Const ERR_EMPLOYEE_REQUIRED As Long = vbObjectError + 1201
Private Const ERR_ASSIGNMENT_REQUIRED As Long = vbObjectError + 1202
Private Const ERR_INVALID_DATE_RANGE As Long = vbObjectError + 1203

Private mConfiguration As clsConfiguracio
Private mEmployees As Collection
Private mAssignments As Collection
Private mStartDate As Date
Private mEndDate As Date
Private mHasPlanningPeriod As Boolean

Private Sub Class_Initialize()
    Set mEmployees = New Collection
    Set mAssignments = New Collection
End Sub

Private Sub Class_Terminate()
    Set mAssignments = Nothing
    Set mEmployees = Nothing
    Set mConfiguration = Nothing
End Sub

Public Property Get Configuration() As clsConfiguracio
    Set Configuration = mConfiguration
End Property

Public Property Set Configuration(ByVal Value As clsConfiguracio)
    If Value Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_CONFIGURATION_REQUIRED, _
            ProcedureName:="Configuration", _
            Description:="La configuració no pot ser Nothing."
    End If

    Set mConfiguration = Value
End Property

Public Property Get Employees() As Collection
    Set Employees = mEmployees
End Property

Public Property Get Assignments() As Collection
    Set Assignments = mAssignments
End Property

Public Property Get StartDate() As Date
    EnsurePlanningPeriodIsDefined
    StartDate = mStartDate
End Property

Public Property Get EndDate() As Date
    EnsurePlanningPeriodIsDefined
    EndDate = mEndDate
End Property

Public Property Get HasPlanningPeriod() As Boolean
    HasPlanningPeriod = mHasPlanningPeriod
End Property

Public Sub SetPlanningPeriod( _
    ByVal PlanningStartDate As Date, _
    ByVal PlanningEndDate As Date)

    If PlanningEndDate < PlanningStartDate Then
        RaiseContextError _
            ErrorNumber:=ERR_INVALID_DATE_RANGE, _
            ProcedureName:="SetPlanningPeriod", _
            Description:="La data final no pot ser anterior a la data inicial."
    End If

    mStartDate = PlanningStartDate
    mEndDate = PlanningEndDate
    mHasPlanningPeriod = True
End Sub

Public Sub AddEmployee(ByVal Employee As clsOperari)
    If Employee Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_EMPLOYEE_REQUIRED, _
            ProcedureName:="AddEmployee", _
            Description:="L'operari no pot ser Nothing."
    End If

    mEmployees.Add Employee
End Sub

Public Sub AddAssignment(ByVal Assignment As clsShiftAssignment)
    If Assignment Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_REQUIRED, _
            ProcedureName:="AddAssignment", _
            Description:="L'assignació no pot ser Nothing."
    End If

    mAssignments.Add Assignment
End Sub

Public Sub ClearEmployees()
    Set mEmployees = New Collection
End Sub

Public Sub ClearAssignments()
    Set mAssignments = New Collection
End Sub

Public Property Get EmployeeCount() As Long
    EmployeeCount = mEmployees.Count
End Property

Public Property Get AssignmentCount() As Long
    AssignmentCount = mAssignments.Count
End Property

Public Function IsReadyForScheduling() As Boolean
    IsReadyForScheduling = _
        Not mConfiguration Is Nothing _
        And mEmployees.Count > 0 _
        And mHasPlanningPeriod
End Function

Private Sub EnsurePlanningPeriodIsDefined()
    If Not mHasPlanningPeriod Then
        RaiseContextError _
            ErrorNumber:=ERR_INVALID_DATE_RANGE, _
            ProcedureName:="EnsurePlanningPeriodIsDefined", _
            Description:="El període de planificació encara no està definit."
    End If
End Sub

Private Sub RaiseContextError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=CLASS_NAME & "." & ProcedureName, _
        Description:=Description
End Sub