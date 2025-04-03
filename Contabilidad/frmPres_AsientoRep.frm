VERSION 5.00
Begin VB.Form frmPres_AsientoRep 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte de Asientos Presupuestarios"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtPeriodo 
      Height          =   315
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   240
      Width           =   4185
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   1410
      MaxLength       =   4
      TabIndex        =   2
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   240
      Width           =   645
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   960
      MaxLength       =   2
      TabIndex        =   1
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   240
      Width           =   435
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ajustes"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmPres_AsientoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReporte_Click()

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ContaExpress"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Mascara='" & vParametros.MascaraCod & "'"
 .Formulas(4) = "SubTitulo='Periodo : " & txtPeriodo & "'"
 .Connect = glogon.ConectRPT
 
 .ReportFileName = App.Path & "\PreAsientos.rpt"

 If Mid(cbo.Text, 1, 1) = "P" Then
    .SelectionFormula = "{PRE_ASIENTOS.COD_CONTABILIDAD} = " & vParametros.CodigoEmpresa _
          & " AND {PRE_ASIENTOS.IANIO} = " & txtAnio _
          & " AND {PRE_ASIENTOS.IMES} = " & txtMes
          
 Else
    .SelectionFormula = "{PRE_ASIENTOS.COD_CONTABILIDAD} = " & vParametros.CodigoEmpresa _
          & " AND {PRE_ASIENTOS.EANIO} = " & txtAnio _
          & " AND {PRE_ASIENTOS.EMES} = " & txtMes
 End If
  .PrintReport
  
End With



End Sub

Private Sub txtMes_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnio, txtMes, txtPeriodo)
End Sub

Private Sub txtAnio_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnio, txtMes, txtPeriodo)
End Sub

Private Sub sbRefrescaInformacion(vAnio As Long, vMes As Integer, Obx As Object)
Dim strResultado As String

On Error GoTo vError
  
  Select Case vMes
    Case 1
        strResultado = "ENERO DEL " & vAnio
    Case 2
        strResultado = "FEBRERO DEL " & vAnio
    Case 3
        strResultado = "MARZO DEL " & vAnio
    Case 4
        strResultado = "ABRIL DEL " & vAnio
    Case 5
        strResultado = "MAYO DEL " & vAnio
    Case 6
        strResultado = "JUNIO DEL " & vAnio
    Case 7
        strResultado = "JULIO DEL " & vAnio
    Case 8
        strResultado = "AGOSTO DEL " & vAnio
    Case 9
        strResultado = "SETIEMBRE DEL " & vAnio
    Case 10
        strResultado = "OCTUBRE DEL " & vAnio
    Case 11
        strResultado = "NOVIEMBRE DEL " & vAnio
    Case 12
        strResultado = "DICIEMBRE DEL " & vAnio
  End Select

  Obx.Text = strResultado

Exit Sub

vError:
End Sub


Private Sub Form_Load()
Dim vFecha As Date

vFecha = fxFechaServidor

txtAnio = Year(vFecha)
txtMes = Month(vFecha)

cbo.AddItem "Positivos"
cbo.AddItem "Negativos"
cbo.Text = "Positivos"

On Error Resume Next
Call sbRefrescaInformacion(txtAnio, txtMes, txtPeriodo)


End Sub
