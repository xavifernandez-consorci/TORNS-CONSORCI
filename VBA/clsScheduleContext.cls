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
' Version:      1.1
'
' Purpose:
'   Encapsulate all data required during schedule generation.
'
' Responsibilities:
'   - Hold the active configuration.
'   - Hold the employee collection.
'   - Define the planning period.
'   - Hold generated assignments.
'   - Index assignments by planning date.
'   - Find, replace and remove assignments safely.
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
Private Const ERR_DUPLICATE_ASSIGNMENT As Long = vbObjectError + 1204
Private Const ERR_ASSIGNMENT_NOT_FOUND As Long = vbObjectError + 1205
Private Const ERR_EMPLOYEE_ID_REQUIRED As Long = vbObjectError + 1206

Private mConfiguration As clsConfiguracio
Private mEmployees As Collection
Private mAssignments As Collection
Private mAssignmentsByDate As Collection

Private mStartDate As Date
Private mEndDate As Date
Private mHasPlanningPeriod As Boolean

'===============================================================================
' Initializes internal collections.
'===============================================================================
Private Sub Class_Initialize()

    Set mEmployees = New Collection
    Set mAssignments = New Collection
    Set mAssignmentsByDate = New Collection

End Sub

'===============================================================================
' Releases object references.
'===============================================================================
Private Sub Class_Terminate()

    Set mAssignmentsByDate = Nothing
    Set mAssignments = Nothing
    Set mEmployees = Nothing
    Set mConfiguration = Nothing

End Sub

'===============================================================================
' Active application configuration.
'===============================================================================
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

'===============================================================================
' Employee collection.
'
' Callers receive the live collection. Employees must be added through
' AddEmployee to preserve validation.
'===============================================================================
Public Property Get Employees() As Collection

    Set Employees = mEmployees

End Property

'===============================================================================
' Assignment collection.
'
' The collection is maintained internally together with the date index.
'===============================================================================
Public Property Get Assignments() As Collection

    Set Assignments = mAssignments

End Property

'===============================================================================
' Planning-period properties.
'===============================================================================
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

'===============================================================================
' Defines the inclusive planning period.
'===============================================================================
Public Sub SetPlanningPeriod( _
    ByVal PlanningStartDate As Date, _
    ByVal PlanningEndDate As Date)

    If PlanningStartDate <= 0 Or PlanningEndDate <= 0 Then
        RaiseContextError _
            ErrorNumber:=ERR_INVALID_DATE_RANGE, _
            ProcedureName:="SetPlanningPeriod", _
            Description:="Les dates del període de planificació no són vàlides."
    End If

    If PlanningEndDate < PlanningStartDate Then
        RaiseContextError _
            ErrorNumber:=ERR_INVALID_DATE_RANGE, _
            ProcedureName:="SetPlanningPeriod", _
            Description:="La data final no pot ser anterior a la data inicial."
    End If

    mStartDate = DateValue(PlanningStartDate)
    mEndDate = DateValue(PlanningEndDate)
    mHasPlanningPeriod = True

End Sub

'===============================================================================
' Adds one validated employee to the context.
'===============================================================================
Public Sub AddEmployee(ByVal Employee As clsOperari)

    If Employee Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_EMPLOYEE_REQUIRED, _
            ProcedureName:="AddEmployee", _
            Description:="L'operari no pot ser Nothing."
    End If

    If Not Employee.IsValid Then
        RaiseContextError _
            ErrorNumber:=ERR_EMPLOYEE_REQUIRED, _
            ProcedureName:="AddEmployee", _
            Description:="L'operari no conté totes les dades obligatòries."
    End If

    mEmployees.Add Employee

End Sub

'===============================================================================
' Adds one assignment and updates the date index.
'
' Only one assignment per employee and date is permitted.
'===============================================================================
Public Sub AddAssignment(ByVal Assignment As clsShiftAssignment)

    Dim assignmentDate As Date
    Dim employeeId As String

    If Assignment Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_REQUIRED, _
            ProcedureName:="AddAssignment", _
            Description:="L'assignació no pot ser Nothing."
    End If

    If Not Assignment.IsValid Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_REQUIRED, _
            ProcedureName:="AddAssignment", _
            Description:="L'assignació no conté totes les dades obligatòries."
    End If

    assignmentDate = Assignment.AssignmentDate
    employeeId = Assignment.Employee.Id

    If AssignmentExists(employeeId, assignmentDate) Then
        RaiseContextError _
            ErrorNumber:=ERR_DUPLICATE_ASSIGNMENT, _
            ProcedureName:="AddAssignment", _
            Description:="Ja existeix una assignació per a aquest operari i data."
    End If

    mAssignments.Add Assignment
    AddAssignmentToDateIndex Assignment

End Sub

'===============================================================================
' Returns all assignments for one date.
'
' An empty collection is returned when the date has no assignments.
'===============================================================================
Public Function GetAssignmentsByDate( _
    ByVal AssignmentDate As Date) As Collection

    Dim result As Collection
    Dim indexedAssignments As Collection
    Dim assignmentItem As Variant

    Set result = New Collection
    Set indexedAssignments = GetIndexedAssignments(AssignmentDate)

    If Not indexedAssignments Is Nothing Then
        For Each assignmentItem In indexedAssignments
            result.Add assignmentItem
        Next assignmentItem
    End If

    Set GetAssignmentsByDate = result

End Function

'===============================================================================
' Finds one assignment by employee identifier and date.
'
' Returns Nothing when no matching assignment exists.
'===============================================================================
Public Function FindAssignment( _
    ByVal EmployeeId As String, _
    ByVal AssignmentDate As Date) As clsShiftAssignment

    Dim indexedAssignments As Collection
    Dim assignmentItem As Variant
    Dim assignment As clsShiftAssignment
    Dim normalizedEmployeeId As String

    normalizedEmployeeId = Trim$(EmployeeId)

    If Len(normalizedEmployeeId) = 0 Then
        RaiseContextError _
            ErrorNumber:=ERR_EMPLOYEE_ID_REQUIRED, _
            ProcedureName:="FindAssignment", _
            Description:="L'identificador de l'operari és obligatori."
    End If

    Set indexedAssignments = GetIndexedAssignments(AssignmentDate)

    If indexedAssignments Is Nothing Then Exit Function

    For Each assignmentItem In indexedAssignments

        Set assignment = assignmentItem

        If StrComp( _
            assignment.Employee.Id, _
            normalizedEmployeeId, _
            vbTextCompare) = 0 Then

            Set FindAssignment = assignment
            Exit Function
        End If

    Next assignmentItem

End Function

'===============================================================================
' Replaces the assignment for the same employee and date.
'
' The replacement assignment must identify an existing employee/date pair.
'===============================================================================
Public Sub ReplaceAssignment(ByVal Assignment As clsShiftAssignment)

    Dim employeeId As String
    Dim assignmentDate As Date

    If Assignment Is Nothing Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_REQUIRED, _
            ProcedureName:="ReplaceAssignment", _
            Description:="L'assignació de substitució no pot ser Nothing."
    End If

    If Not Assignment.IsValid Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_REQUIRED, _
            ProcedureName:="ReplaceAssignment", _
            Description:="L'assignació de substitució no és vàlida."
    End If

    employeeId = Assignment.Employee.Id
    assignmentDate = Assignment.AssignmentDate

    If Not AssignmentExists(employeeId, assignmentDate) Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_NOT_FOUND, _
            ProcedureName:="ReplaceAssignment", _
            Description:="No existeix cap assignació per substituir."
    End If

    RemoveAssignment employeeId, assignmentDate
    AddAssignment Assignment

End Sub

'===============================================================================
' Removes one assignment by employee identifier and date.
'===============================================================================
Public Sub RemoveAssignment( _
    ByVal EmployeeId As String, _
    ByVal AssignmentDate As Date)

    Dim normalizedEmployeeId As String
    Dim currentAssignments As Collection
    Dim assignmentItem As Variant
    Dim assignment As clsShiftAssignment
    Dim wasRemoved As Boolean

    normalizedEmployeeId = Trim$(EmployeeId)

    If Len(normalizedEmployeeId) = 0 Then
        RaiseContextError _
            ErrorNumber:=ERR_EMPLOYEE_ID_REQUIRED, _
            ProcedureName:="RemoveAssignment", _
            Description:="L'identificador de l'operari és obligatori."
    End If

    Set currentAssignments = New Collection

    For Each assignmentItem In mAssignments

        Set assignment = assignmentItem

        If StrComp( _
            assignment.Employee.Id, _
            normalizedEmployeeId, _
            vbTextCompare) = 0 _
            And assignment.AssignmentDate = AssignmentDate Then

            wasRemoved = True
        Else
            currentAssignments.Add assignment
        End If

    Next assignmentItem

    If Not wasRemoved Then
        RaiseContextError _
            ErrorNumber:=ERR_ASSIGNMENT_NOT_FOUND, _
            ProcedureName:="RemoveAssignment", _
            Description:="No s'ha trobat l'assignació que es vol eliminar."
    End If

    Set mAssignments = currentAssignments
    RebuildAssignmentDateIndex

End Sub

'===============================================================================
' Clears all employees from the context.
'===============================================================================
Public Sub ClearEmployees()

    Set mEmployees = New Collection

End Sub

'===============================================================================
' Clears all generated in-memory assignments and their index.
'===============================================================================
Public Sub ClearAssignments()

    Set mAssignments = New Collection
    Set mAssignmentsByDate = New Collection

End Sub

'===============================================================================
' Collection counters.
'===============================================================================
Public Property Get EmployeeCount() As Long

    EmployeeCount = mEmployees.Count

End Property

Public Property Get AssignmentCount() As Long

    AssignmentCount = mAssignments.Count

End Property

'===============================================================================
' Indicates whether the minimum information required by the motor exists.
'===============================================================================
Public Function IsReadyForScheduling() As Boolean

    IsReadyForScheduling = _
        Not mConfiguration Is Nothing _
        And mEmployees.Count > 0 _
        And mHasPlanningPeriod

End Function

'===============================================================================
' Returns True when an employee already has an assignment for the date.
'===============================================================================
Private Function AssignmentExists( _
    ByVal EmployeeId As String, _
    ByVal AssignmentDate As Date) As Boolean

    Dim existingAssignment As clsShiftAssignment

    Set existingAssignment = FindAssignment(EmployeeId, AssignmentDate)
    AssignmentExists = Not existingAssignment Is Nothing

End Function

'===============================================================================
' Adds an assignment to its indexed date bucket.
'===============================================================================
Private Sub AddAssignmentToDateIndex( _
    ByVal Assignment As clsShiftAssignment)

    Dim dateKey As String
    Dim indexedAssignments As Collection

    dateKey = BuildDateKey(Assignment.AssignmentDate)
    Set indexedAssignments = GetIndexedAssignments(Assignment.AssignmentDate)

    If indexedAssignments Is Nothing Then
        Set indexedAssignments = New Collection
        mAssignmentsByDate.Add indexedAssignments, dateKey
    End If

    indexedAssignments.Add Assignment

End Sub

'===============================================================================
' Returns the indexed collection for a date.
'
' Returns Nothing when the date has no indexed assignments.
'===============================================================================
Private Function GetIndexedAssignments( _
    ByVal AssignmentDate As Date) As Collection

    Dim dateKey As String

    dateKey = BuildDateKey(AssignmentDate)

    On Error Resume Next
    Set GetIndexedAssignments = mAssignmentsByDate.Item(dateKey)
    On Error GoTo 0

End Function

'===============================================================================
' Rebuilds the complete date index from the assignment collection.
'===============================================================================
Private Sub RebuildAssignmentDateIndex()

    Dim assignmentItem As Variant
    Dim assignment As clsShiftAssignment

    Set mAssignmentsByDate = New Collection

    For Each assignmentItem In mAssignments
        Set assignment = assignmentItem
        AddAssignmentToDateIndex assignment
    Next assignmentItem

End Sub

'===============================================================================
' Builds a locale-independent key for one calendar date.
'===============================================================================
Private Function BuildDateKey( _
    ByVal AssignmentDate As Date) As String

    BuildDateKey = Format$(DateValue(AssignmentDate), "yyyymmdd")

End Function

'===============================================================================
' Ensures that the planning period has been defined.
'===============================================================================
Private Sub EnsurePlanningPeriodIsDefined()

    If Not mHasPlanningPeriod Then
        RaiseContextError _
            ErrorNumber:=ERR_INVALID_DATE_RANGE, _
            ProcedureName:="EnsurePlanningPeriodIsDefined", _
            Description:="El període de planificació encara no està definit."
    End If

End Sub

'===============================================================================
' Raises a controlled domain-layer error.
'===============================================================================
Private Sub RaiseContextError( _
    ByVal ErrorNumber As Long, _
    ByVal ProcedureName As String, _
    ByVal Description As String)

    Err.Raise _
        Number:=ErrorNumber, _
        Source:=CLASS_NAME & "." & ProcedureName, _
        Description:=Description

End Sub