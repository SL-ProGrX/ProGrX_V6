VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Begin VB.Form frmCO_ControlCartasAvisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de Cartas de Avisos"
   ClientHeight    =   7200
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10728
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCO_ControlCartasAvisos.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   10728
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5172
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   10452
      _Version        =   1245186
      _ExtentX        =   18436
      _ExtentY        =   9123
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   312
      Left            =   9480
      TabIndex        =   11
      Top             =   240
      Width           =   1092
      _Version        =   1245186
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Buscar"
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
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   372
      Index           =   0
      Left            =   3480
      TabIndex        =   6
      Top             =   6240
      Width           =   3612
      _Version        =   1245186
      _ExtentX        =   6371
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Notificar: Primer Aviso"
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
      Value           =   -1  'True
   End
   Begin VB.CheckBox chkMarcas 
      Appearance      =   0  'Flat
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   240
      Width           =   2175
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   312
      Left            =   1560
      TabIndex        =   5
      Top             =   6240
      Width           =   1332
      _Version        =   1245186
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   372
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   6600
      Width           =   3612
      _Version        =   1245186
      _ExtentX        =   6371
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Notificar: Segundo Aviso"
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
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   492
      Left            =   7320
      TabIndex        =   8
      Top             =   6360
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Picture         =   "frmCO_ControlCartasAvisos.frx":169B2
   End
   Begin XtremeSuiteControls.PushButton btnMail 
      Height          =   492
      Left            =   8880
      TabIndex        =   9
      Top             =   6360
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Email"
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
      Picture         =   "frmCO_ControlCartasAvisos.frx":1716E
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "para pago de mora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   1560
      TabIndex        =   3
      Top             =   6600
      Width           =   1452
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   2
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "frmCO_ControlCartasAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub sbAplica()
Dim strSQL  As String, x As Long

If txtUsuario = "" Then Exit Sub

Me.MousePointer = vbHourglass

For x = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(x).Checked Then
    
'    strSQL = "exec spCBRControlAsg '" & lsw.ListItems.Item(x).SubItems(1) _
'           & "','" &  txt txtTrasladar.Text & "',1"
'    Call ConectionExecute(strSQL)
    
   ' Call Bitacora("Aplica", "Traslado Caso CBR Ced:" & lsw.ListItems.Item(x).SubItems(1) & " de " & txtCodigo _
             & " a " & txtTrasladar)
 End If
Next x

Me.MousePointer = vbDefault

MsgBox "Envio de Avisos realizado Satisfactoriamente...", vbInformation
Call sbBuscar


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnMail_Click()

MsgBox "Los Casos han sido enviados a la bandeja de notificaciones por Email!", vbInformation

End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
chkMarcas.Value = vbUnchecked

strSQL = "select * from vCBRControlListado" _
       & " where Usuario = '" & txtUsuario & "' and Mora > 0" _
       & " and cedula not in(select cedula from cbr_seguimiento" _
       & " where dateadd(d,tiempo_resolucion,fecha) >= dbo.MyGetdate() and cod_gestion in(select valor" _
       & " from cbr_parametros where cod_parametro in('03','04') ) )"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!fecha_asignacion, "dd/mm/yyyy"))
     itmX.SubItems(1) = rs!Cedula
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = IIf((rs!mantener = 1), "SI", "NO")
     itmX.SubItems(4) = Format(rs!Mora, "Standard")
     itmX.SubItems(5) = rs!CuotaMora
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnReporte_Click()
Call sbAplica
End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()

vModulo = 4

With lsw.ColumnHeaders
    .Add , , "Fecha", 1500
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3100
    .Add , , "Mantiene?", 1500, vbCenter
    .Add , , "Mora Actual", 2100
    .Add , , "No Cuota", 2100, vbCenter
End With

dtpVence.Value = fxFechaServidor

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub chkMarcas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkMarcas.Value
Next i

End Sub





Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then optX.Item(0).SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "usuario"
    gBusquedas.Orden = "usuario"
    gBusquedas.Filtro = " and estado = 1"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub txtUsuarioDesc_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkMantener.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = " and estado = 1"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If

End Sub


