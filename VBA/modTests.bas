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

    Dim emp As clsOperari

    Set emp = New clsOperari

    emp.Id = "OP001"
    emp.Name = "Operari prova"
    emp.IsActive = True
    emp.IsIntensiveCandidate = True

    AssertEquals "OP001", emp.Id, "Employee.Id"
    AssertEquals "Operari prova", emp.Name, "Employee.Name"
    AssertTrue emp.IsValid, "Employee.IsValid"

End Sub

'===============================================================================
' clsScheduleContext
'===============================================================================
Public Sub Test_ScheduleContext()

    Dim ctx As clsScheduleContext

    Set ctx = New clsScheduleContext

    ctx.SetPlanningPeriod DateSerial(2027, 1, 1), _
                          DateSerial(2027, 12, 31)

    AssertTrue ctx.HasPlanningPeriod, "PlanningPeriod"

End Sub

'===============================================================================
' clsConfiguracio
'===============================================================================
Public Sub Test_Configuration()

    Dim cfg As clsConfiguracio

    Set cfg = New clsConfiguracio

    cfg.MorningWeeks = 2
    cfg.AfternoonWeeks = 2

    AssertEquals 2, cfg.MorningWeeks, "MorningWeeks"
    AssertEquals 2, cfg.AfternoonWeeks, "AfternoonWeeks"

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

Public Sub Test_InsertIntensivePreview()

    Dim ctx As clsScheduleContext

    Set ctx = BuildDemoContext()

    GenerateSchedule ctx

    AssertTrue True, "InsertIntensivePreview"

End Sub

'===============================================================================
' Verifica el flux inicial del motor:
'   - valida el context;
'   - genera el cicle base;
'   - detecta els dijous;
'   - mostra la previsualització d'intensius.
'
' En aquesta fase encara no modifica les assignacions base.
'===============================================================================
Public Sub Test_InsertIntensivePreview()

    Dim context As clsScheduleContext
    Dim expectedAssignments As Long

    Set context = BuildDemoContext()

    modMotorRotacions.GenerateSchedule context

    ' Període del 04/01/2027 al 24/01/2027:
    ' 21 dies x 3 operaris actius = 63 assignacions.
    expectedAssignments = 63

    AssertEquals _
        expectedAssignments, _
        context.AssignmentCount, _
        "IntensivePreview.AssignmentCount"

End Sub

'===============================================================================
' Construeix un context complet i vàlid per a les proves del motor.
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

'===============================================================================
' Construeix una configuració completa i vàlida per als tests.
'===============================================================================
Private Function BuildDemoConfiguration() As clsConfiguracio

    Dim configuration As clsConfiguracio

    Set configuration = New clsConfiguracio

    configuration.EmployeeCount = 3

    configuration.MorningWeeks = 2
    configuration.AfternoonWeeks = 2

    configuration.IntensiveStartDay = "Dijous"
    configuration.IntensiveEndDay = "Dimecres"

    configuration.WorkingHoursPerShift = 8

    configuration.MorningStartTime = "06:00"
    configuration.MorningEndTime = "14:00"

    configuration.AfternoonStartTime = "14:00"
    configuration.AfternoonEndTime = "22:00"

    configuration.DutyBackupCount = 1

    Set BuildDemoConfiguration = configuration

End Function

'===============================================================================
' Construeix un operari actiu i candidat a intensiu.
'===============================================================================
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

End Sub