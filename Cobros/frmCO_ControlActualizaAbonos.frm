VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Begin VB.Form frmCO_ControlActualizaAbonos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de Abonos"
   ClientHeight    =   2736
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10632
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2736
   ScaleWidth      =   10632
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   168
      Left            =   0
      TabIndex        =   0
      Top             =   2568
      Width           =   10632
      _ExtentX        =   18754
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdActualizar 
      Height          =   732
      Left            =   8040
      TabIndex        =   2
      Top             =   1440
      Width           =   1692
      _Version        =   1245186
      _ExtentX        =   2984
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Actualizar"
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
      Appearance      =   14
      Picture         =   "frmCO_ControlActualizaAbonos.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculo de recaudos por gestion de cobros!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   7212
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCO_ControlActualizaAbonos.frx":06FC
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_ControlActualizaAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualizar_Click()
Dim strSQL As String


On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Caption = "Actualizando ....(Espere)"

strSQL = "exec spCBRControlSGTAbonoGeneral"
Call ConectionExecute(strSQL)

lbl.Caption = "Actualización de Abonos realizada satisfactoriamente..."

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


