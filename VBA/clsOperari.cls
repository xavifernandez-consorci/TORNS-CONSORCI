Option Explicit

' Purpose: Represents a single employee in the scheduling application.
' Responsibilities:
' - Encapsulate employee identity and state.
' - Validate assignment of values at the object boundary.
' - Provide safe defaults for new instances.

Private mId As String
Private mName As String
Private mIsActive As Boolean
Private mIsDutyBackup As Boolean
Private mIsIntensiveCandidate As Boolean

Private mRotationState As String
Private mIntensiveCount As Long
Private mLastIntensiveDate As Date
Private mHasLastIntensiveDate As Boolean

' Purpose: Initializes the object with safe default values.
Private Sub Class_Initialize()
    InitializeDefaults
End Sub

' Purpose: Resets the employee values to their default state.
Public Sub InitializeDefaults()
    mId = vbNullString
    mName = vbNullString
    mIsActive = True
    mIsDutyBackup = False
    mIsIntensiveCandidate = False
End Sub

' Purpose: Returns the employee identifier.
Public Property Get Id() As String
    Id = mId
End Property

Public Property Get RotationState() As String
    RotationState = mRotationState
End Property

Public Property Let RotationState(ByVal Value As String)
    mRotationState = Trim$(Value)
End Property

Public Property Get IntensiveCount() As Long
    IntensiveCount = mIntensiveCount
End Property

Public Property Let IntensiveCount(ByVal Value As Long)
    If Value < 0 Then
        Err.Raise vbObjectError + 1500, _
                  "clsOperari.IntensiveCount", _
                  "El nombre d'intensius no pot ser negatiu."
    End If

    mIntensiveCount = Value
End Property

Public Property Get HasLastIntensiveDate() As Boolean
    HasLastIntensiveDate = mHasLastIntensiveDate
End Property

Public Property Get LastIntensiveDate() As Date

    If Not mHasLastIntensiveDate Then
        Err.Raise vbObjectError + 1501, _
                  "clsOperari.LastIntensiveDate", _
                  "L'operari encara no té cap intensiu registrat."
    End If

    LastIntensiveDate = mLastIntensiveDate

End Property

Public Property Let LastIntensiveDate(ByVal Value As Date)

    mLastIntensiveDate = Value
    mHasLastIntensiveDate = True

End Property

' Purpose: Sets the employee identifier.
' Parameters:
' - value: The identifier to assign.
' Raises:
' - An error when the identifier is empty or exceeds the allowed length.
Public Property Let Id(ByVal value As String)
    Dim trimmedValue As String
    trimmedValue = Trim$(value)

    If Len(trimmedValue) = 0 Then
        Err.Raise vbObjectError + 513, "clsOperari.Id", "L'identificador de l'operari és obligatori."
    End If

    If Len(trimmedValue) > 50 Then
        Err.Raise vbObjectError + 514, "clsOperari.Id", "L'identificador de l'operari no pot superar 50 caràcters."
    End If

    If InStr(trimmedValue, vbCr) > 0 Or InStr(trimmedValue, vbLf) > 0 Or InStr(trimmedValue, vbTab) > 0 Then
        Err.Raise vbObjectError + 515, "clsOperari.Id", "L'identificador de l'operari no pot contenir salts de línia ni tabulacions."
    End If

    mId = trimmedValue
End Property

' Purpose: Returns the employee display name.
Public Property Get Name() As String
    Name = mName
End Property

' Purpose: Sets the employee display name.
' Parameters:
' - value: The display name to assign.
' Raises:
' - An error when the name is empty or exceeds the allowed length.
Public Property Let Name(ByVal value As String)
    Dim trimmedValue As String
    trimmedValue = Trim$(value)

    If Len(trimmedValue) = 0 Then
        Err.Raise vbObjectError + 516, "clsOperari.Name", "El nom de l'operari és obligatori."
    End If

    If Len(trimmedValue) > 100 Then
        Err.Raise vbObjectError + 517, "clsOperari.Name", "El nom de l'operari no pot superar 100 caràcters."
    End If

    mName = trimmedValue
End Property

' Purpose: Returns whether the employee is currently active.
Public Property Get IsActive() As Boolean
    IsActive = mIsActive
End Property

' Purpose: Sets whether the employee is currently active.
' Parameters:
' - value: The active state to assign.
Public Property Let IsActive(ByVal value As Boolean)
    mIsActive = value
End Property

' Purpose: Returns whether the employee is a duty backup operator.
Public Property Get IsDutyBackup() As Boolean
    IsDutyBackup = mIsDutyBackup
End Property

' Purpose: Sets whether the employee is a duty backup operator.
' Parameters:
' - value: The backup state to assign.
Public Property Let IsDutyBackup(ByVal value As Boolean)
    mIsDutyBackup = value
End Property

' Purpose: Returns whether the employee is eligible for intensive rotation.
Public Property Get IsIntensiveCandidate() As Boolean
    IsIntensiveCandidate = mIsIntensiveCandidate
End Property

' Purpose: Sets whether the employee is eligible for intensive rotation.
' Parameters:
' - value: The intensive eligibility state to assign.
Public Property Let IsIntensiveCandidate(ByVal value As Boolean)
    mIsIntensiveCandidate = value
End Property

'===============================================================================
' Creates a deep copy of the current employee.
'===============================================================================
Public Function Clone() As clsOperari

    Dim employee As clsOperari

    Set employee = New clsOperari

    employee.Id = Me.Id
    employee.Name = Me.Name
    employee.IsActive = Me.IsActive
    employee.IsDutyBackup = Me.IsDutyBackup
    employee.IsIntensiveCandidate = Me.IsIntensiveCandidate

    employee.RotationState = Me.RotationState
    employee.IntensiveCount = Me.IntensiveCount

    If Me.HasLastIntensiveDate Then
        employee.LastIntensiveDate = Me.LastIntensiveDate
    End If

    Set Clone = employee

End Function

'===============================================================================
' Indicates whether the object contains a valid employee.
'===============================================================================
Public Function IsValid() As Boolean

    IsValid = _
        Len(mId) > 0 _
        And Len(mName) > 0

End Function
