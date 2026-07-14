Attribute VB_Name = "modTests"
Option Explicit

Private mTestsRun As Long
Private mTestsPassed As Long

'===============================================================================
' Entry point
'===============================================================================
Public Sub RunAllTests()

    mTestsRun = 0
    mTestsPassed = 0

    Debug.Print String(70, "=")
    Debug.Print "TORNS CONSORCI - TEST SUITE"
    Debug.Print String(70, "=")

    Test_EmployeeCreation
    Test_ScheduleContext
    Test_Configuration
    Test_InsertIntensivePreview
    Test_ManualOverrideMetadata
    Test_ReplaceIntensiveBlock

    Debug.Print String(70, "-")
    Debug.Print "Tests executats : "; mTestsRun
    Debug.Print "Tests correctes : "; mTestsPassed
    Debug.Print "Tests fallits   : "; mTestsRun - mTestsPassed
    Debug.Print String(70, "=")

End Sub

'===============================================================================
' clsOperari
'===============================================================================
Public Sub Test_EmployeeCreation()

    Dim employee As clsOperari

    Set employee = New clsOperari

    employee.Id = "OP001"
    employee.Name = "Operari prova"
    employee.IsActive = True
    employee.IsIntensiveCandidate = True

    AssertEquals "OP001", employee.Id, "Employee.Id"
    AssertEquals "Operari prova", employee.Name, "Employee.Name"
    AssertTrue employee.IsValid, "Employee.IsValid"

End Sub

'===============================================================================
' clsScheduleContext
'===============================================================================
Public Sub Test_ScheduleContext()

    Dim context As clsScheduleContext

    Set context = New clsScheduleContext

    context.SetPlanningPeriod _
        PlanningStartDate:=DateSerial(2027, 1, 1), _
        PlanningEndDate:=DateSerial(2027, 12, 31)

    AssertTrue context.HasPlanningPeriod, "PlanningPeriod"

End Sub

'===============================================================================
' clsConfiguracio
'===============================================================================
Public Sub Test_Configuration()

    Dim configuration As clsConfiguracio

    Set configuration = New clsConfiguracio

    configuration.MorningWeeks = 2
    configuration.AfternoonWeeks = 2

    AssertEquals 2, configuration.MorningWeeks, "MorningWeeks"
    AssertEquals 2, configuration.AfternoonWeeks, "AfternoonWeeks"

End Sub

'===============================================================================
' modMotorRotacions
'===============================================================================
Public Sub Test_InsertIntensivePreview()

    Dim context As clsScheduleContext
    Dim expectedAssignments As Long

    Set context = BuildDemoContext()

    modMotorRotacions.GenerateSchedule context

    ' 21 dies x 3 operaris actius = 63 assignacions.
    expectedAssignments = 63

    AssertEquals _
        expectedAssignments, _
        context.AssignmentCount, _
        "IntensivePreview.AssignmentCount"

End Sub

'===============================================================================
' clsShiftAssignment
'===============================================================================
Public Sub Test_ManualOverrideMetadata()

    Dim originalEmployee As clsOperari
    Dim replacementEmployee As clsOperari
    Dim assignment As clsShiftAssignment

    Set originalEmployee = BuildDemoEmployee( _
        EmployeeId:="OP001", _
        EmployeeName:="Operari original")

    Set replacementEmployee = BuildDemoEmployee( _
        EmployeeId:="OP002", _
        EmployeeName:="Operari substitut")

    Set assignment = New clsShiftAssignment

    Set assignment.Employee = replacementEmployee
    assignment.AssignmentDate = DateSerial(2027, 1, 7)
    assignment.ShiftCode = "I"
    assignment.IsIntensive = True
    assignment.IsPrimaryDuty = True

    assignment.MarkAsManualOverride _
        OriginalEmployeeId:=originalEmployee.Id, _
        OverrideReason:="Baixa medica"

    AssertTrue _
        assignment.IsManualOverride, _
        "ManualOverride.IsManualOverride"

    AssertEquals _
        originalEmployee.Id, _
        assignment.OriginalEmployeeId, _
        "ManualOverride.OriginalEmployeeId"

    AssertEquals _
        "Baixa medica", _
        assignment.OverrideReason, _
        "ManualOverride.OverrideReason"

    AssertTrue _
        assignment.IsValid, _
        "ManualOverride.IsValid"

End Sub

'===============================================================================
' modSubstitucions
'
' Cas provat:
'   - OP001 estava planificat per a l'intensiu del dijous 07/01/2027.
'   - OP001 causa baixa des del dilluns 04/01/2027.
'   - OP002 el substitueix.
'
' Resultat esperat:
'   - OP001: baixa de dilluns a diumenge.
'   - OP002: descans de dilluns a dimecres.
'   - OP002: intensiu de dijous a diumenge.
'   - Dilluns 11/01/2027 d'OP002 queda intacte.
'   - L'historial d'intensius d'OP002 augmenta en una unitat.
'===============================================================================
Public Sub Test_ReplaceIntensiveBlock()

    Dim context As clsScheduleContext
    Dim originalEmployee As clsOperari
    Dim replacementEmployee As clsOperari

    Dim intensiveStartDate As Date
    Dim weekMonday As Date
    Dim followingMonday As Date

    Dim dayOffset As Long
    Dim assignmentDate As Date

    Dim replacementCountBefore As Long
    Dim followingMondayShiftBefore As String
    Dim followingMondayAssignment As clsShiftAssignment

    Set context = BuildDemoContext()

    modMotorRotacions.GenerateSchedule context

    Set originalEmployee = context.Employees(1)
    Set replacementEmployee = context.Employees(2)

    intensiveStartDate = DateSerial(2027, 1, 7)
    weekMonday = DateSerial(2027, 1, 4)
    followingMonday = DateSerial(2027, 1, 11)

    replacementCountBefore = replacementEmployee.IntensiveCount

    Set followingMondayAssignment = context.FindAssignment( _
        EmployeeId:=replacementEmployee.Id, _
        AssignmentDate:=followingMonday)

    followingMondayShiftBefore = followingMondayAssignment.ShiftCode

    modSubstitucions.ReplaceIntensiveBlock _
        Context:=context, _
        OriginalEmployee:=originalEmployee, _
        ReplacementEmployee:=replacementEmployee, _
        IntensiveStartDate:=intensiveStartDate, _
        Reason:="Baixa medica"

    ' Operari original: baixa de dilluns a diumenge.
    For dayOffset = 0 To 6

        assignmentDate = DateAdd("d", dayOffset, weekMonday)

        AssertAssignmentShift _
            Context:=context, _
            EmployeeId:=originalEmployee.Id, _
            AssignmentDate:=assignmentDate, _
            ExpectedShiftCode:="B", _
            TestName:="Substitution.Original.Baixa." & CStr(dayOffset + 1)

    Next dayOffset

    ' Operari substitut: descans de dilluns a dimecres.
    For dayOffset = 0 To 2

        assignmentDate = DateAdd("d", dayOffset, weekMonday)

        AssertAssignmentShift _
            Context:=context, _
            EmployeeId:=replacementEmployee.Id, _
            AssignmentDate:=assignmentDate, _
            ExpectedShiftCode:="D", _
            TestName:="Substitution.Replacement.Descans." & CStr(dayOffset + 1)

    Next dayOffset

    ' Operari substitut: intensiu de dijous a diumenge.
    For dayOffset = 3 To 6

        assignmentDate = DateAdd("d", dayOffset, weekMonday)

        AssertAssignmentShift _
            Context:=context, _
            EmployeeId:=replacementEmployee.Id, _
            AssignmentDate:=assignmentDate, _
            ExpectedShiftCode:="I", _
            TestName:="Substitution.Replacement.Intensiu." & CStr(dayOffset - 2)

    Next dayOffset

    Set followingMondayAssignment = context.FindAssignment( _
        EmployeeId:=replacementEmployee.Id, _
        AssignmentDate:=followingMonday)

    AssertEquals _
        followingMondayShiftBefore, _
        followingMondayAssignment.ShiftCode, _
        "Substitution.FollowingMonday.Unchanged"

    AssertEquals _
        replacementCountBefore + 1, _
        replacementEmployee.IntensiveCount, _
        "Substitution.Replacement.IntensiveCount"

    AssertAssignmentManualOverride _
        Context:=context, _
        EmployeeId:=replacementEmployee.Id, _
        AssignmentDate:=intensiveStartDate, _
        OriginalEmployeeId:=originalEmployee.Id, _
        ExpectedReason:="Baixa medica", _
        TestName:="Substitution.ManualOverride"

End Sub

'===============================================================================
' Test-data builders
'===============================================================================
Private Function BuildDemoContext() As clsScheduleContext

    Dim context As clsScheduleContext
    Dim configuration As clsConfiguracio

    Set context = New clsScheduleContext
    Set configuration = BuildDemoConfiguration()

    Set context.Configuration = configuration

    context.SetPlanningPeriod _
        PlanningStartDate:=DateSerial(2027, 1, 4), _
        PlanningEndDate:=DateSerial(2027, 1, 24)

    context.AddEmployee BuildDemoEmployee( _
        EmployeeId:="OP001", _
        EmployeeName:="Operari 1")

    context.AddEmployee BuildDemoEmployee( _
        EmployeeId:="OP002", _
        EmployeeName:="Operari 2")

    context.AddEmployee BuildDemoEmployee( _
        EmployeeId:="OP003", _
        EmployeeName:="Operari 3")

    Set BuildDemoContext = context

End Function

Private Function BuildDemoConfiguration() As clsConfiguracio

    Dim configuration As clsConfiguracio

    Set configuration = New clsConfiguracio

    configuration.EmployeeCount = 3
    configuration.MorningWeeks = 2
    configuration.AfternoonWeeks = 2
    configuration.IntensiveStartDay = "Dijous"
    configuration.IntensiveEndDay = "Diumenge"
    configuration.WorkingHoursPerShift = 8
    configuration.MorningStartTime = "06:00"
    configuration.MorningEndTime = "14:00"
    configuration.AfternoonStartTime = "14:00"
    configuration.AfternoonEndTime = "22:00"
    configuration.DutyBackupCount = 1

    Set BuildDemoConfiguration = configuration

End Function

Private Function BuildDemoEmployee( _
    ByVal EmployeeId As String, _
    ByVal EmployeeName As String) As clsOperari

    Dim employee As clsOperari

    Set employee = New clsOperari

    employee.Id = EmployeeId
    employee.Name = EmployeeName
    employee.IsActive = True
    employee.IsIntensiveCandidate = True
    employee.IsDutyBackup = False
    employee.IntensiveCount = 0
    employee.RotationState = "MATI"

    Set BuildDemoEmployee = employee

End Function

'===============================================================================
' Test helpers
'===============================================================================
Private Sub AssertAssignmentShift( _
    ByVal Context As clsScheduleContext, _
    ByVal EmployeeId As String, _
    ByVal AssignmentDate As Date, _
    ByVal ExpectedShiftCode As String, _
    ByVal TestName As String)

    Dim assignment As clsShiftAssignment

    Set assignment = Context.FindAssignment( _
        EmployeeId:=EmployeeId, _
        AssignmentDate:=AssignmentDate)

    AssertTrue Not assignment Is Nothing, TestName & ".Exists"

    If Not assignment Is Nothing Then
        AssertEquals _
            ExpectedShiftCode, _
            assignment.ShiftCode, _
            TestName & ".ShiftCode"
    End If

End Sub

Private Sub AssertAssignmentManualOverride( _
    ByVal Context As clsScheduleContext, _
    ByVal EmployeeId As String, _
    ByVal AssignmentDate As Date, _
    ByVal OriginalEmployeeId As String, _
    ByVal ExpectedReason As String, _
    ByVal TestName As String)

    Dim assignment As clsShiftAssignment

    Set assignment = Context.FindAssignment( _
        EmployeeId:=EmployeeId, _
        AssignmentDate:=AssignmentDate)

    AssertTrue Not assignment Is Nothing, TestName & ".Exists"

    If Not assignment Is Nothing Then

        AssertTrue _
            assignment.IsManualOverride, _
            TestName & ".IsManualOverride"

        AssertEquals _
            OriginalEmployeeId, _
            assignment.OriginalEmployeeId, _
            TestName & ".OriginalEmployeeId"

        AssertEquals _
            ExpectedReason, _
            assignment.OverrideReason, _
            TestName & ".OverrideReason"

    End If

End Sub

'===============================================================================
' Assertions
'===============================================================================
Private Sub AssertTrue( _
    ByVal Condition As Boolean, _
    ByVal TestName As String)

    mTestsRun = mTestsRun + 1

    If Condition Then
        mTestsPassed = mTestsPassed + 1
        Debug.Print "[OK]   "; TestName
    Else
        Debug.Print "[FAIL] "; TestName
    End If

End Sub

Private Sub AssertEquals( _
    ByVal Expected As Variant, _
    ByVal Actual As Variant, _
    ByVal TestName As String)

    AssertTrue Expected = Actual, TestName

End Sub
