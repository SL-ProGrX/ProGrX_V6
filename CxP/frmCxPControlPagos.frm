VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxPControlPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Pagos a Proveedores"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   11325
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5052
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   8911
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Detalle"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Resumen"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswResumen"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4572
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   10812
         _Version        =   1441793
         _ExtentX        =   19071
         _ExtentY        =   8064
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswResumen 
         Height          =   4572
         Left            =   -69880
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   10812
         _Version        =   1441793
         _ExtentX        =   19071
         _ExtentY        =   8064
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   732
      Left            =   8880
      TabIndex        =   9
      Top             =   2400
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   16777215
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
      Appearance      =   17
      Picture         =   "frmCxPControlPagos.frx":0000
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   732
      Left            =   9960
      TabIndex        =   10
      Top             =   2400
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Informe"
      BackColor       =   16777215
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
      Appearance      =   17
      Picture         =   "frmCxPControlPagos.frx":0A1E
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.ComboBox cboCancela 
      Height          =   312
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   6840
      TabIndex        =   14
      Top             =   1680
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFecha 
      Height          =   312
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkProvTodos 
      Height          =   264
      Left            =   9000
      TabIndex        =   16
      Top             =   1320
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Todos?"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2400
      TabIndex        =   19
      Top             =   2760
      Width           =   1332
      _Version        =   1441793
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   3720
      TabIndex        =   20
      Top             =   2760
      Width           =   1332
      _Version        =   1441793
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1440
      TabIndex        =   21
      Top             =   1320
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2520
      TabIndex        =   22
      Top             =   1320
      Width           =   6252
      _Version        =   1441793
      _ExtentX        =   11028
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFactura 
      Height          =   312
      Left            =   1440
      TabIndex        =   23
      Top             =   1680
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   315
      Left            =   5160
      TabIndex        =   24
      Top             =   2760
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNSolicitud 
      Height          =   312
      Left            =   7200
      TabIndex        =   25
      Top             =   2760
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Pagos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   6852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   5040
      TabIndex        =   7
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No.Solicitud"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   7200
      TabIndex        =   5
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Base"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancela"
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
      TabIndex        =   0
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11412
   End
End
Attribute VB_Name = "frmCxPControlPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vInicia As Boolean

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboFecha.SetFocus
vError:
End Sub

Private Sub cboFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus
End Sub

Private Sub chkProvTodos_Click()
If chkProvTodos.Value = vbChecked Then
 txtCodigo.Locked = False
Else
 txtCodigo.Locked = True
End If
End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency
Dim vColumna As String

On Error GoTo vError

Me.MousePointer = vbHourglass


Select Case Mid(cboFecha, 1, 1)
  Case "E" 'Fecha de Envio a Tesoreria
    vColumna = "C.Fecha_Traslada"
  
  Case "V" 'Fecha de Vencimiento
    vColumna = "C.Fecha_Vencimiento"
    
  Case "C" 'Fecha de Emisión
    vColumna = "T.Fecha_Emision"
  
  Case Else
    vColumna = "C.Fecha_Vencimiento"
End Select

strSQL = "select C.*,P.descripcion as Proveedor,B.descripcion as 'Banco', T.ndocumento" _
       & " from cxp_PagoProv C inner join CxP_Proveedores P on P.cod_proveedor = C.cod_proveedor" _
       & " left join Tes_Transacciones T on C.tesoreria = T.nsolicitud" _
       & " left join Tes_Bancos B on T.id_banco = B.id_Banco" _
       & " where " & vColumna & " between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " and C.Tipo_Cancelacion = '" & Mid(cboCancela.Text, 1, 1) & "'"

If chkProvTodos.Value = vbUnchecked And IsNumeric(txtCodigo) Then
   strSQL = strSQL & " and C.cod_proveedor = " & txtCodigo
End If


Select Case Mid(cboEstado.Text, 1, 1)
  Case "S" 'Sin enviar a tesoreria
     strSQL = strSQL & " and C.tesoreria is null"
  Case "E" 'Todas las Enviadas
     strSQL = strSQL & " and C.tesoreria is not null"
  Case "P" 'Pendientes
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO = 'P'"
  Case "C" 'Canceladas
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO in('I','T','E')"
  Case "A" 'Anuladas
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO in('A','N')"
     
  Case "T" 'Todas
End Select


If Len(Trim(txtFactura.Text)) > 0 Then
   strSQL = strSQL & " and C.cod_factura like '%" & txtFactura.Text & "%'"
End If

If Len(Trim(txtDocumento.Text)) > 0 Then
   strSQL = strSQL & " and T.Ndocumento like '%" & txtDocumento.Text & "%'"
End If

If Len(Trim(txtNSolicitud.Text)) > 0 Then
   strSQL = strSQL & " and C.Tesoreria = " & txtNSolicitud.Text
End If



lsw.ListItems.Clear
curTotal = 0
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_Factura)
     itmX.SubItems(1) = rs!Npago
     itmX.SubItems(2) = rs!cod_proveedor
     itmX.SubItems(3) = Format(rs!Fecha_Vencimiento, "yyyy/mm/dd")
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Proveedor
     itmX.SubItems(6) = IIf(IsNull(rs!tesoreria), 0, rs!tesoreria)
     itmX.SubItems(7) = rs!Banco & "" 'Banco
     itmX.SubItems(8) = rs!nDocumento & "" 'Documento
 curTotal = curTotal + rs!Monto
 rs.MoveNext
Loop
rs.Close

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(4) = "________"
Set itmX = lsw.ListItems.Add(, , "Total")
    itmX.SubItems(4) = Format(curTotal, "Standard")
    itmX.Bold = True
    itmX.ForeColor = vbBlue

'Resumen
strSQL = "select count(*) as Pagos, sum(C.monto) as Monto,P.descripcion as Proveedor, P.cod_proveedor " _
       & " from cxp_PagoProv C inner join CxP_Proveedores P on P.cod_proveedor = C.cod_proveedor" _
       & " left join Tes_Transacciones T on C.tesoreria = T.nsolicitud" _
       & " left join Tes_Bancos B on T.id_banco = B.id_Banco" _
       & " where " & vColumna & " between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " and C.Tipo_Cancelacion = '" & Mid(cboCancela.Text, 1, 1) & "'"
        
If chkProvTodos.Value = vbUnchecked And IsNumeric(txtCodigo) Then
   strSQL = strSQL & " and C.cod_proveedor = " & txtCodigo
End If

Select Case Mid(cboEstado.Text, 1, 1)
  Case "S" 'Sin enviar a tesoreria
     strSQL = strSQL & " and C.tesoreria is null"
  Case "E" 'Todas las Enviadas
     strSQL = strSQL & " and C.tesoreria is not null"
  Case "P" 'Pendientes
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO = 'P'"
  Case "C" 'Canceladas
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO in('I','T','E')"
  Case "A" 'Anuladas
     strSQL = strSQL & " and C.tesoreria is not null and C.TESORERIA_ESTADO in('A','N')"
     
  Case "T" 'Todas
End Select


If Len(Trim(txtFactura.Text)) > 0 Then
   strSQL = strSQL & " and C.cod_factura like '%" & txtFactura.Text & "%'"
End If

If Len(Trim(txtDocumento.Text)) > 0 Then
   strSQL = strSQL & " and T.Ndocumento like '%" & txtDocumento.Text & "%'"
End If

If Len(Trim(txtNSolicitud.Text)) > 0 Then
   strSQL = strSQL & " and C.Tesoreria = " & txtNSolicitud.Text
End If


strSQL = strSQL & " group by P.cod_proveedor,P.descripcion"



lswResumen.ListItems.Clear

curTotal = 0
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswResumen.ListItems.Add(, , rs!cod_proveedor)
     itmX.SubItems(1) = rs!Proveedor
     itmX.SubItems(2) = rs!pagos
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
 curTotal = curTotal + rs!Monto
 rs.MoveNext
Loop
rs.Close

Set itmX = lswResumen.ListItems.Add(, , "")
    itmX.SubItems(3) = "________"
Set itmX = lswResumen.ListItems.Add(, , "Total")
    itmX.SubItems(3) = Format(curTotal, "Standard")
    itmX.Bold = True
    itmX.ForeColor = vbBlue


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporte()
Dim vSQL As String, vSubTitulo As String
Dim vColumna As String

On Error GoTo vError

Me.MousePointer = vbHourglass


Select Case Mid(cboFecha, 1, 1)
  Case "E" 'Fecha de Envio a Tesoreria
    vColumna = "Fecha_Traslada"
  
  Case "V" 'Fecha de Vencimiento
    vColumna = "Fecha_Vencimiento"
'
'  Case "C" 'Fecha de Emisión
'    vColumna = "Fecha_Emision"
  
  Case Else
    vColumna = "Fecha_Vencimiento"
End Select
vColumna = UCase(vColumna)

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes Cuentas x Pagar"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 

vSQL = "CDATE({CXP_PAGOPROV." & vColumna & "}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") _
     & ") AND {CXP_PAGOPROV.TIPO_CANCELACION} = '" & Mid(cboCancela.Text, 1, 1) & "'"


vSubTitulo = cboFecha.Text & " ...: [Rango.: " & Format(dtpInicio.Value, "dd/mm/yyyy") & " - " & Format(dtpCorte.Value, "dd/mm/yyyy") & "] CANCELADO VÍA.:" & UCase(cboCancela.Text)
                
If chkProvTodos.Value = vbUnchecked And IsNumeric(txtCodigo) Then
   vSQL = vSQL & " AND {CXP_PROVEEDORES.COD_PROVEEDOR} = " & txtCodigo
  vSubTitulo = vSubTitulo & " / PROVEEDOR COD." & txtCodigo
Else
  vSubTitulo = vSubTitulo & " / TODOS LOS PROVEEDORES"
End If

Select Case Mid(cboEstado.Text, 1, 1)
  Case "S" 'Sin Enviar a Tesoreria
     vSQL = vSQL & " AND ISNULL({CXP_PAGOPROV.TESORERIA}) = TRUE"
  Case "E" 'Todas las enviadas
     vSQL = vSQL & " AND ISNULL({CXP_PAGOPROV.TESORERIA}) = FALSE"
  Case "C" 'Canceladas
     vSQL = vSQL & " AND ISNULL({CXP_PAGOPROV.TESORERIA}) = FALSE and ({CXP_PAGOPROV.TESORERIA_ESTADO} = 'I' or {CXP_PAGOPROV.TESORERIA_ESTADO} = 'T' or {CXP_PAGOPROV.TESORERIA_ESTADO} = 'E')"
  Case "P" 'Pendiente de Pago
     vSQL = vSQL & " AND ISNULL({CXP_PAGOPROV.TESORERIA}) = FALSE and ({CXP_PAGOPROV.TESORERIA_ESTADO} = 'P')"
  Case "A" 'Anuladas
     vSQL = vSQL & " AND ISNULL({CXP_PAGOPROV.TESORERIA}) = FALSE and ({CXP_PAGOPROV.TESORERIA_ESTADO} = 'A' or {CXP_PAGOPROV.TESORERIA_ESTADO} = 'N'"

End Select
     vSubTitulo = vSubTitulo & " / " & UCase(cboEstado.Text)
 
    .Formulas(3) = "fxTitulo = 'PROGRAMACION DE PAGOS'"
    .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("CxP_ProgramacionListadoDetalle.rpt")
    .SelectionFormula = vSQL
 
 .PrintReport

End With
Me.MousePointer = vbDefault


Exit Sub

vError:

Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture


With lsw.ColumnHeaders
   .Clear
   .Add , , "No. Factura", 2400
   .Add , , "No. Pago", 1000, vbCenter
   .Add , , "Prov.Id.", 1000, vbCenter
   .Add , , "Vence", 1200, vbCenter
   .Add , , "Monto", 1300, vbRightJustify
   .Add , , "Proveedor", 4000
   .Add , , "No.Tesorería", 1200, vbCenter
   .Add , , "Cuenta Bancos", 3000
   .Add , , "No.Documento", 2000, vbCenter
End With

With lswResumen.ColumnHeaders
   .Clear
   .Add , , "Prov.Id.", 1200, vbCenter
   .Add , , "Proveedor", 4000
   .Add , , "Qty/Pagos", 1200, vbCenter
   .Add , , "Monto", 1600, vbRightJustify
End With

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cboCancela.Clear
cboCancela.AddItem "Desembolso"
cboCancela.AddItem "Cargos (Ajuste)"
cboCancela.Text = "Desembolso"

cboEstado.Clear
cboEstado.AddItem "Sin Enviar a Tesoreria"
cboEstado.AddItem "Enviadas a Tesoreria"
cboEstado.AddItem "Pendientes (Pago)"
cboEstado.AddItem "Canceladas (Pago)"
cboEstado.AddItem "Anuladas (Pago)"
cboEstado.AddItem "Todas"
cboEstado.Text = "Todas"


cboFecha.Clear
cboFecha.AddItem "Vencimiento"
cboFecha.AddItem "Envío a Tesorería"
cboFecha.AddItem "Cancelación"
cboFecha.Text = "Envío a Tesorería"

vInicia = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub






Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
       Call sbBuscar
  Case "Reporte"
       Call sbReporte
End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado3
  
  lsw.ListItems.Clear
  Call sbBuscar
End If

End Sub

Private Sub txtCodigo_LostFocus()
txtNombre.Text = fxSIFCCodigos("D", txtCodigo, "proveedores")
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNSolicitud.SetFocus
End Sub

Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus

If KeyCode = vbKeyF4 And chkProvTodos.Value = xtpUnchecked And IsNumeric(txtCodigo.Text) Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select cod_factura,total,fecha From vCxP_ProgramacionPago"
  gBusquedas.Filtro = " and CxP_Estado = 'G' and Cod_Proveedor = " & txtCodigo
  frmBusquedas.Show vbModal
  txtFactura.Text = gBusquedas.Resultado
End If


End Sub

Private Sub txtNSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbBuscar
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFactura.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado3
  
  lsw.ListItems.Clear
  Call sbBuscar
End If

End Sub

