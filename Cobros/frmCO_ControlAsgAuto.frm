VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCO_ControlAsgAuto 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación Automática de Casos"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswGrupos 
      Height          =   1575
      Left            =   1680
      TabIndex        =   13
      Top             =   1800
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   2778
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbOpciones 
      Height          =   975
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9970
      _ExtentY        =   1714
      _StockProps     =   79
      Caption         =   "Casos a procesa: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbOpcion 
         Height          =   264
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Casos con cuentas con Morosidad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbOpcion 
         Height          =   264
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Casos con cuentas al Día"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.CheckBox chkMantenerNuevos 
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Marcar como mantener nuevos casos asignados."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   5652
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdAplica 
      Height          =   735
      Left            =   7680
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCO_ControlAsgAuto.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.GroupBox gbOpciones 
      Height          =   1455
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   4800
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9970
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Tipo de Ingreso a la Lista: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbOpcion 
         Height          =   264
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Asignar únicamente casos nuevos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbOpcion 
         Height          =   264
         Index           =   3
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Inicializar la Lista de Asignación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkInicializaTodo 
         Height          =   612
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Todos los Casos (Morosos / Al Día)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Grupo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lbl 
      Height          =   852
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   5412
      _Version        =   1441793
      _ExtentX        =   9546
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   $"frmCO_ControlAsgAuto.frx":09C3
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1440
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_ControlAsgAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTipo_Click()
If cboTipo.ListCount = 0 Then Exit Sub

If cboTipo.Text = "Distribución por Roles de Usuario" Then
    gbOpciones.Item(0).Enabled = False
Else
    gbOpciones.Item(0).Enabled = True
End If

End Sub

Private Sub cmdAplica_Click()
Dim strSQL As String, vInicializa As Integer, rs As New ADODB.Recordset


strSQL = "select usuario from cbr_usuarios where estado = 1"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No existen usuarios activos para realizar asignacion, verifique...", vbExclamation
  Exit Sub
End If
rs.Close

On Error GoTo vError


Me.MousePointer = vbHourglass


lbl.Caption = "Procesando Torneo de Asignación de Casos de Cobros a Oficiales..."
DoEvents

'Inicializacion
Select Case True
   Case rbOpcion.Item(2).Value 'Nuevos Casos
           vInicializa = 0
   Case rbOpcion.Item(3).Value 'Inicializa
       If chkInicializaTodo.Value = vbChecked Then
           vInicializa = 1 'Inicializa TODO
       Else
           vInicializa = 2 'Inicializa por Tipo de Lista
       End If
End Select


'Casos a Procesar
If cboTipo.Text = "Distribución por Roles de Usuario" Then
    
    Select Case True
       Case rbOpcion.Item(0).Value 'Morosos
            strSQL = "exec spCBRControlDistribucion 'R'," & vInicializa & "," & chkMantenerNuevos.Value & ",1,0"
       Case rbOpcion.Item(1).Value 'Al Día
            strSQL = "exec spCBRControlDistribucion 'R'," & vInicializa & "," & chkMantenerNuevos.Value & ",0,1"
    End Select


Else
    'Distribución STANDARD
    Select Case True
       Case rbOpcion.Item(0).Value 'Morosos
            strSQL = "exec spCBRControlDistribucion 'S'," & vInicializa & "," & chkMantenerNuevos.Value & ",1,0"
       Case rbOpcion.Item(1).Value 'Al Día
            strSQL = "exec spCBRControlDistribucion 'S'," & vInicializa & "," & chkMantenerNuevos.Value & ",0,1"
    End Select

End If

Call ConectionExecute(strSQL)

lbl.Caption = "Proceso Finalizado Satisfactoriamente..."

Me.MousePointer = vbDefault

MsgBox "Proceso Terminado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 4


Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

cboTipo.Clear
cboTipo.AddItem "Distribución Standard"
cboTipo.AddItem "Distribución por Roles de Usuario"
cboTipo.Text = "Distribución por Roles de Usuario"



With lswGrupos.ColumnHeaders
    .Clear
    .Add , , , lswGrupos.Width - 250
End With
lswGrupos.HideColumnHeaders = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


