VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPosConsecutivos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consecutivos: Sistema Comercial"
   ClientHeight    =   4764
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4764
   ScaleWidth      =   8784
   Begin VB.TextBox txtDP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   22
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtDevoluciones 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   20
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtRecibos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtNC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtOrdenSalida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   15
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtPedidos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtFacturaManual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtFacturaAuto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtOrdenEntrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtTraslados 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtSalidas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtEntradas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   540
      Left            =   6600
      TabIndex        =   24
      Top             =   3960
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmPosConsecutivos.frx":0000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Consecutivos para Documentos Comerciales: Inventarios y Puntos de ventas"
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
      Height          =   852
      Left            =   1800
      TabIndex        =   25
      Top             =   240
      Width           =   6732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   11
      Left            =   4440
      TabIndex        =   23
      Top             =   3120
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Devoluciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recibos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   360
      TabIndex        =   19
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas de Crédito"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   4440
      TabIndex        =   18
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos / Apartados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas Manuales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas Automáticas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenes de Salidas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenes de Entradas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Traslados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salidas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entradas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13092
   End
End
Attribute VB_Name = "frmPosConsecutivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGuardar_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "update pv_consecutivos set entradas = " & txtEntradas _
       & ",salidas = " & txtSalidas & ",traslados = " & txtTraslados _
       & ",facturas_auto = " & txtFacturaAuto & ",facturas_man = " & txtFacturaManual _
       & ",orden_entrada = " & txtOrdenEntrada & ",orden_salida = " & txtOrdenSalida _
       & ",pedidos = " & txtPedidos & ",recibos = " & txtRecibos & ",nc  = " & txtNC _
       & ",devolucion = " & txtDevoluciones & ",depositos = " & txtDP
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Consecutivos del Sistema Comercial")

Me.MousePointer = vbDefault
MsgBox "Información Guardada Satisfactorimente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 34
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 34

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select * from pv_consecutivos"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 strSQL = "insert pv_consecutivos(entradas,salidas,traslados,facturas_auto" _
         & ",facturas_man,pedidos,orden_entrada,orden_salida,recibos,nc,devolucion,depositos)" _
         & " values(0,0,0,0,0,0,0,0,0,0,0,0)"
  Call ConectionExecute(strSQL)
  txtEntradas = 0
  txtSalidas = 0
  txtTraslados = 0
  txtOrdenEntrada = 0
  txtOrdenSalida = 0
  txtFacturaAuto = 0
  txtFacturaManual = 0
  txtPedidos = 0
  txtRecibos = 0
  txtNC = 0
  txtDevoluciones = 0
  txtDP = 0
Else
  txtEntradas = rs!entradas
  txtSalidas = rs!salidas
  txtTraslados = rs!traslados
  txtOrdenEntrada = rs!orden_entrada
  txtOrdenSalida = rs!orden_salida
  txtFacturaAuto = rs!facturas_auto
  txtFacturaManual = rs!facturas_man
  txtPedidos = rs!pedidos
  txtRecibos = rs!Recibos
  txtNC = rs!nc
  txtDevoluciones = rs!devolucion
  txtDP = rs!depositos
End If
rs.Close


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
