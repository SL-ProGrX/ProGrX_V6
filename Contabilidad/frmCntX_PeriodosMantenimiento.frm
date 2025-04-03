VERSION 5.00
Begin VB.Form frmCntX_Periodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento - Periodos"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   HelpContextID   =   2005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtHasta 
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Año de Corte del Periodo Fiscal"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtDesde 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Año de Inicio del Periodo Fiscal"
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo Fiscal a Definir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmCntX_Periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicar_Click()
Dim strSQL As String, vAnio As Long, vMes As Integer
Dim blnPaso As Boolean

On Error GoTo vError

vAnio = txtDesde
vMes = 10
blnPaso = True

Do While blnPaso
    strSQL = "insert into periodos(cod_empresa,anio,mes,estado) values(" _
           & vCodEmpresa & "," & vAnio & "," & vMes & ",'P')"
    glogon.Conection.Execute strSQL
    
    If vMes = 12 Then
       vMes = 1
       vAnio = vAnio + 1
    Else
       vMes = vMes + 1
    End If
    
    
    If Val(txtDesde) + 1 And vMes = 10 Then blnPaso = False
    
Loop


MsgBox "Periodo Fiscal Creado...", vbInformation

vError:

End Sub

Private Sub cmdReporte_Click()
Call sbReportes("PERIODOS", Me)
End Sub

Private Sub Form_Load()
Set Me.Icon = frmContenedor.Icon
txtDesde = Year(fxFechaServidor)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtDesde_Change()
On Error GoTo vError
 
 txtHasta = Val(txtDesde) + 1

vError:
End Sub
