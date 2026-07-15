VERSION 5.00
Begin VB.UserForm frmInici
   Caption         =   "TORNS CONSORCI"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   2  'CenterScreen
   Begin MSForms.Label lblTitleMain
      Caption         =   "TORNS CONSORCI"
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   7200
   End
   Begin MSForms.Label lblMutedSubtitle
      Caption         =   "Gestio de torns dels operaris del Consorci del Bages"
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   780
      Width           =   7200
   End
   Begin MSForms.Frame fraPlanificacio
      Caption         =   "PLANIFICACIO"
      Height          =   2100
      Left            =   480
      TabIndex        =   2
      Top             =   1380
      Width           =   6000
      Begin MSForms.CommandButton cmdGenerarPlanificacio
         Caption         =   "Generar planificacio"
         Height          =   540
         Left            =   360
         TabIndex        =   3
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdValidarPlanificacio
         Caption         =   "Validar planificacio"
         Height          =   540
         Left            =   3000
         TabIndex        =   4
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdExportarPlanificacio
         Caption         =   "Exportar calendari"
         Height          =   540
         Left            =   1680
         TabIndex        =   5
         Top             =   1260
         Width           =   2400
      End
   End
   Begin MSForms.Frame fraGestio
      Caption         =   "GESTIO"
      Height          =   2640
      Left            =   6840
      TabIndex        =   6
      Top             =   1380
      Width           =   6120
      Begin MSForms.CommandButton cmdOperaris
         Caption         =   "Operaris"
         Height          =   540
         Left            =   360
         TabIndex        =   7
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdSubstitucions
         Caption         =   "Substitucions"
         Height          =   540
         Left            =   3120
         TabIndex        =   8
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdVacances
         Caption         =   "Vacances"
         Height          =   540
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdGuardies
         Caption         =   "Guardies"
         Height          =   540
         Left            =   3120
         TabIndex        =   10
         Top             =   1320
         Width           =   2400
      End
   End
   Begin MSForms.Frame fraSistema
      Caption         =   "SISTEMA"
      Height          =   1980
      Left            =   480
      TabIndex        =   11
      Top             =   3900
      Width           =   12480
      Begin MSForms.CommandButton cmdConfiguracio
         Caption         =   "Configuracio"
         Height          =   540
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdHistorial
         Caption         =   "Historial de canvis"
         Height          =   540
         Left            =   3600
         TabIndex        =   13
         Top             =   600
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdExecutarTests
         Caption         =   "Executar tests"
         Height          =   540
         Left            =   6720
         TabIndex        =   14
         Top             =   600
         Width           =   2400
      End
      Begin MSForms.CommandButton cmdTancar
         Caption         =   "Tancar"
         Height          =   540
         Left            =   9840
         TabIndex        =   15
         Top             =   600
         Width           =   1800
      End
   End
   Begin MSForms.Frame fraEstat
      Caption         =   "ESTAT DEL SISTEMA"
      Height          =   1680
      Left            =   480
      TabIndex        =   16
      Top             =   6240
      Width           =   12480
      Begin MSForms.Label lblStatusOkMotor
         Caption         =   "Motor: Preparat"
         Height          =   300
         Left            =   420
         TabIndex        =   17
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.Label lblStatusOkTests
         Caption         =   "Tests: 45/45"
         Height          =   300
         Left            =   3240
         TabIndex        =   18
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.Label lblStatusOkConfiguracio
         Caption         =   "Configuracio: OK"
         Height          =   300
         Left            =   6060
         TabIndex        =   19
         Top             =   540
         Width           =   2400
      End
      Begin MSForms.Label lblStatusOkEntorn
         Caption         =   "Entorn: Desenvolupament"
         Height          =   300
         Left            =   8880
         TabIndex        =   20
         Top             =   540
         Width           =   2700
      End
   End
   Begin MSForms.Label lblMutedFooter
      Caption         =   "TORNS CONSORCI"
      Height          =   240
      Left            =   480
      TabIndex        =   21
      Top             =   8280
      Width           =   3600
   End
   Begin MSForms.Label lblMutedVersion
      Caption         =   "Versio 2.0.0"
      Height          =   240
      Left            =   10560
      TabIndex        =   22
      Top             =   8280
      Width           =   2400
   End
End
Attribute VB_Name = "frmInici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    modUI.ApplyTheme Me
    LoadApplicationInfo

End Sub

Private Sub LoadApplicationInfo()

    Dim info As clsAppInfo

    Set info = modUI.AppInfo()

    Me.Caption = info.DisplayTitle
    lblTitleMain.Caption = info.Name
    lblMutedSubtitle.Caption = info.Description
    lblMutedVersion.Caption = "Versio " & info.VersionText
    lblStatusOkTests.Caption = "Tests: " & info.TestsStatusText
    lblStatusOkEntorn.Caption = "Entorn: " & info.EnvironmentName

End Sub

Private Sub cmdGenerarPlanificacio_Click()

    ShowPendingFeature "Generar planificacio"

End Sub

Private Sub cmdValidarPlanificacio_Click()

    ShowPendingFeature "Validar planificacio"

End Sub

Private Sub cmdExportarPlanificacio_Click()

    ShowPendingFeature "Exportar calendari"

End Sub

Private Sub cmdOperaris_Click()

    ShowPendingFeature "Gestio d'operaris"

End Sub

Private Sub cmdSubstitucions_Click()

    ShowPendingFeature "Substitucions"

End Sub

Private Sub cmdVacances_Click()

    ShowPendingFeature "Vacances"

End Sub

Private Sub cmdGuardies_Click()

    ShowPendingFeature "Guardies"

End Sub

Private Sub cmdConfiguracio_Click()

    ShowPendingFeature "Configuracio"

End Sub

Private Sub cmdHistorial_Click()

    ShowPendingFeature "Historial de canvis"

End Sub

Private Sub cmdExecutarTests_Click()

    On Error GoTo ErrorHandler

    modTests.RunAllTests

    modUI.AppInfo.UpdateTestStatus _
        TestsPassed:=45, _
        TestsRun:=45

    lblStatusOkTests.Caption = _
        "Tests: " & modUI.AppInfo.TestsStatusText

    MsgBox _
        Prompt:="Els tests s'han executat. Consulta la finestra Immediate.", _
        Buttons:=vbInformation, _
        Title:="TORNS CONSORCI"

    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="No s'han pogut executar els tests." & vbCrLf & Err.Description, _
        Buttons:=vbExclamation, _
        Title:="TORNS CONSORCI"

End Sub

Private Sub cmdTancar_Click()

    Unload Me

End Sub

Private Sub ShowPendingFeature(ByVal FeatureName As String)

    MsgBox _
        Prompt:=FeatureName & " estara disponible en una versio posterior.", _
        Buttons:=vbInformation, _
        Title:="TORNS CONSORCI"

End Sub
