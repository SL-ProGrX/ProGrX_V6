VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCierreMensualASE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre Manual de Auxiliares de Cuentas Corrientes"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   HelpContextID   =   7001
   Icon            =   "frmCierreMensualASE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdCierre 
      Height          =   852
      Left            =   7320
      TabIndex        =   0
      Top             =   1440
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Cierre Manual de Auxiliares"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCierreMensualASE.frx":08CA
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierre de Auxiliares"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   7212
   End
   Begin XtremeSuiteControls.Label lbl 
      Height          =   1332
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   2350
      _StockProps     =   79
      Caption         =   "..."
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
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCierreMensualASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCierre_Click()
Dim strSQL As String, i As Byte
Dim vFecha As Date

'PASOS
'1. Guardar El Estado Actual de los Creditos
'   saldo_inicial,saldo,proceso,opex,id_solicitud,codigo
'2. Actualizar Total_debitos y Total_creditos con los
'   movimientos del mes
'3. Establecer Nuevo Corte de Saldos.
'4. Insertar en Historicos, el periodo procesado.
'5. Crear Referencia Contable (Metodo Contable)


i = MsgBox("Esta seguro que desea establecer cierre del Mes y Nuevo saldo inicial, se le recuerda" _
                   & " que tiene que ser el ultimo día del mes, cuando ya no se procese información", vbYesNo)
If i = vbNo Then Exit Sub


lbl.Alignment = 0

lbl.Caption = "Actualizando Auxiliares (Espere...)" & vbCrLf & " - Esta operación puede durar varios minutos..."

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

strSQL = "exec spSIFAuxMain " & Year(vFecha) & "," & Month(vFecha) & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
lbl.Caption = "Cierre Concluido Satisfactoriamente...."

Exit Sub


vError:
  lbl.Caption = "Error...."
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub Form_Load()

vModulo = 10 'Cuentas Corrientes
Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

lbl.Caption = "Este proceso realiza el cierre del mes del los auxiliares (Patrimonio, Excedentes, Créditos, Ahorros); el cual crea una copia para uso historico" _
            & " de los Saldos Iniciales y Finales, con sus movimientos respectivos" _
            & " por cortes mensuales" & vbCrLf & vbCrLf _
            & "POR ESTE MOTIVO ESTE PROCESO SOLO DEBE DE SER REALIZADO EL ULTIMO DIA DEL MES (UNA SOLA VEZ)"

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

