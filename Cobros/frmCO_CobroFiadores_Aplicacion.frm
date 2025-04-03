VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCO_CobroFiadores_Aplicacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobro de Fiadores: Aplicación de Abonos a Deuda"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnExpedientes 
      Height          =   330
      Left            =   4800
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _Version        =   1572864
      _ExtentX        =   794
      _ExtentY        =   582
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCO_CobroFiadores_Aplicacion.frx":0000
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13996
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13996
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnProcesar 
         Height          =   495
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Procesar"
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
         Picture         =   "frmCO_CobroFiadores_Aplicacion.frx":0707
         ImageAlignment  =   0
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtExpedientes 
      Height          =   330
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   5175
      _Version        =   1572864
      _ExtentX        =   9128
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aplicacion de Abonos a Deudores"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Total de Casos Pendientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblStatus 
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Este proceso aplica a la Operación del Deudor, los cobros realizados a los fiadores vinculados con la cuenta."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_CobroFiadores_Aplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnProcesar_Click()
Dim i As Integer


i = MsgBox("Esta seguro que desea procesar Abonos desde Cobro a Fiadores?", vbYesNo)
If i = vbNo Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCBR_Cobro_Fiador_Aplica_Abonos 0, '', 0"
Call OpenRecordSet(rs, strSQL)

txtExpedientes.Text = rs!Pendientes

ProgressBarX.Max = rs!Pendientes
ProgressBarX.Value = 0

Do While rs!Pendientes > 0
    strSQL = "exec spCBR_Cobro_Fiador_Aplica_Abonos 20, '" & glogon.Usuario & "', 1"
    Call OpenRecordSet(rs, strSQL)
    ProgressBarX.Value = ProgressBarX.Max - rs!Pendientes
Loop

Me.MousePointer = vbDefault

MsgBox "Proceso de aplicación de Abonos desde cobro a Fiadores realizado satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCasos_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCBR_Cobro_Fiador_Aplica_Abonos 0, '', 0"
Call OpenRecordSet(rs, strSQL)

txtExpedientes.Text = rs!Pendientes

ProgressBarX.Max = rs!Pendientes
ProgressBarX.Value = 0

rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

lblStatus.BackColor = RGB(254, 249, 231) 'Amarillo
'lblStatus.BackColor = RGB(232, 246, 243)  'Verde

Call sbCasos_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
