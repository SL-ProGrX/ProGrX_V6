VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_Facturas_Cancela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de Facturas"
   ClientHeight    =   6630
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9990
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2772
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9732
      _Version        =   1441793
      _ExtentX        =   17166
      _ExtentY        =   4890
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8760
      Top             =   120
   End
   Begin XtremeSuiteControls.FlatEdit feCliente 
      Height          =   312
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feClienteId 
      Height          =   312
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPagador 
      Height          =   312
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboDivisas 
      Height          =   312
      Left            =   8400
      TabIndex        =   6
      Top             =   960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   255
      Index           =   0
      Left            =   9600
      TabIndex        =   7
      ToolTipText     =   "Exportar a Excel"
      Top             =   4200
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCxC_Facturas_Cancela.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      BackColor       =   14737632
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   2778
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   312
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   15
         Top             =   600
         Width           =   5412
         _Version        =   1441793
         _ExtentX        =   9546
         _ExtentY        =   1397
         _StockProps     =   77
         ForeColor       =   0
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   0
         Left            =   6720
         TabIndex        =   16
         Top             =   480
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Pago"
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
         Appearance      =   17
         Picture         =   "frmCxC_Facturas_Cancela.frx":08D1
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   1
         Left            =   7680
         TabIndex        =   17
         Top             =   480
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Appearance      =   17
         Picture         =   "frmCxC_Facturas_Cancela.frx":0D7E
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   2
         Left            =   8520
         TabIndex        =   18
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         Appearance      =   17
         Picture         =   "frmCxC_Facturas_Cancela.frx":1556
         TextImageRelation=   1
      End
      Begin VB.Label Label3 
         Caption         =   "Documento ..:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Notas ..:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Total ..:"
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
         Index           =   4
         Left            =   4080
         TabIndex        =   19
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Label Label27 
      Caption         =   "Diferencia ...:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   11
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "Total a Pagar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente...:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   7
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Id...:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCxC_Facturas_Cancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mClienteId As String, mOperacion As Long, mFactura As String
Dim vPaso As Boolean, pCharRelleno As String

Private Sub sbConsultaCliente()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 
strSQL = "select cedula,nombre from cxc_personas where cedula = '" & mClienteId & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  
    ModuloCajas.mClienteId = Trim(rs!Cedula)
    ModuloCajas.mCliente = Trim(rs!Nombre)
    
    ModuloCajas.mDivisa = "COL" 'RTrim(rs!Divisa)
    ModuloCajas.mConceptoValida = True 'IIf((rs!Caja_Valida_Concepto > 0), True, False)
 
    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
     
    feClienteId.Text = rs!Cedula
    feCliente.Text = rs!Nombre
       
    'Carga Pagadores Con Facturas Pendientes de Pago
    strSQL = "select Per.Cedula as 'IdX', Per.Nombre as 'ItmX'" _
           & " from vCxC_Facturas_Pendientes_Cancelacion Ft inner join CxC_Personas Per on Ft.Cedula_Pagador = Per.Cedula" _
           & " where  Ft.cedula = '" & feClienteId.Text & "'" _
           & " group by Per.Cedula, Per.Nombre"
     vPaso = True
         Call sbCbo_Llena_New(cboPagador, strSQL, False, True)
     vPaso = False
     Call cboPagador_Click
     
Else
 
 MsgBox "No se encontró Cliente?", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbLimpiaDatos()
 feClienteId.Text = ""
 feCliente.Text = ""
  
 txtDiferencia.Text = 0

 txtTotalCajas.Text = 0

 txtNotas.Text = ""
 
 lsw.ListItems.Clear
 lsw.ColumnHeaders.Clear
 lsw.ColumnHeaders.Add , , "No. Operación", 1200
 lsw.ColumnHeaders.Add , , "No. Factura", 2100
 lsw.ColumnHeaders.Add , , "Monto", 1800, vbRightJustify
 lsw.ColumnHeaders.Add , , "Fecha Pago", 1200
 lsw.ColumnHeaders.Add , , "Divisa", 1100, vbCenter
 lsw.ColumnHeaders.Add , , "Importe", 1400, vbRightJustify
 lsw.ColumnHeaders.Add , , "Fecha Emite", 1200
 lsw.ColumnHeaders.Add , , "Fecha Activa", 1200
  
 cboPagador.Clear
 
End Sub


Private Sub btnExport_Click(Index As Integer)
Call Excel_Exportar_Lsw(lsw)
End Sub

Private Sub cboDivisas_Click()
If vPaso Or cboDivisas.ListCount <= 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
    
    
    'Carga Facturas
    strSQL = "select * from vCxC_Facturas_Pendientes_Cancelacion" _
           & " where  cedula = '" & feClienteId.Text & "'" _
           & " and Cedula_Pagador = '" & cboPagador.ItemData(cboPagador.ListIndex) & "'" _
           & " and cod_divisa = '" & cboDivisas.ItemData(cboDivisas.ListIndex) & "'" _
           & " order by CEDULA, FECHA_PAGO, COD_FACTURA"
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Operacion)
          itmX.SubItems(1) = rs!cod_Factura
          itmX.SubItems(2) = Format(rs!Monto, "Standard")
          itmX.SubItems(3) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
          itmX.SubItems(4) = rs!cod_Divisa
          itmX.SubItems(5) = Format(rs!Importe, "Standard")
          itmX.SubItems(6) = Format(rs!Fecha_Emision, "dd/mm/yyyy")
          itmX.SubItems(7) = Format(rs!Activa_Fecha, "dd/mm/yyyy")
      rs.MoveNext
    Loop

End Sub

Private Sub cboPagador_Click()
If vPaso Or cboPagador.ListCount <= 0 Then Exit Sub

Dim strSQL As String

    'Re Inicia Calculos de Cajas
    ModuloCajas.mTiquete = Mid(Trim(feClienteId.Text), 1, 5) & "." & Mid(cboPagador.ItemData(cboPagador.ListIndex), 1, 8) & "." & Format(Time, "HH:mm:ss")
    
    'Carga Divisas de Facturas Pendientes con el Pagador
    strSQL = "select Cod_Divisa as 'IdX', Cod_Divisa as 'ItmX'" _
           & " from vCxC_Facturas_Pendientes_Cancelacion" _
           & " where  cedula = '" & feClienteId.Text & "' and cedula_pagador = '" & cboPagador.ItemData(cboPagador.ListIndex) & "'" _
           & " group by Cod_Divisa"
     
     lsw.ListItems.Clear
     
     vPaso = True
         Call sbCbo_Llena_New(cboDivisas, strSQL, False, True)
     vPaso = False
     
     Call cboDivisas_Click


End Sub

Private Sub Form_Activate()
 vModulo = 31

End Sub

Private Sub Form_Load()

 vModulo = 31
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 Call sbLimpiaDatos
  
If IsNumeric(GLOBALES.gTag) Then
   mOperacion = GLOBALES.gTag
   mFactura = GLOBALES.gTag2
   mClienteId = GLOBALES.gTag3
Else
   mClienteId = ""
   mFactura = ""
   mOperacion = 0
End If

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If Item.Checked Then
    txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text) + CCur(Item.SubItems(5)), "Standard")
Else
    txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text) - CCur(Item.SubItems(5)), "Standard")
End If


txtDiferencia.Text = Format(CCur(txtTotalCajas.Text) - CCur(txtTotalPagar), "Standard")

vError:

End Sub


Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial
End Sub


Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Abonos a Cuentas por Cobrar    ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX' , rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"


If mClienteId <> "" Then
    feClienteId.Text = mClienteId
    Call sbConsultaCliente
End If



End Sub


Private Sub sbDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String, pTipoCambio As Currency
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency

'vCuenta = pCuenta
'
'pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)
'
'
''Cuentas
'strSQL = "exec spCxC_OperacionCtas " & txtOperacion.Text
'Call OpenRecordSet(rs, strSQL)
'
'
'strSQL = "exec spCxC_DocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pNumDoc & "','R'"
'Call OpenRecordSet(rsTmp, strSQL, 0)
'If rsTmp.EOF And rsTmp.BOF Then
'  curIntC = 0
'  curIntM = 0
'  curAmortiza = 0
'  curCargo = 0
'Else
'  curIntC = rsTmp!IntCor
'  curIntM = rsTmp!IntMor
'  curAmortiza = rsTmp!Principal
'  curCargo = rsTmp!Cargos
'End If
'rsTmp.Close
'
'
'
''Lineas de Comprobante
'strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(lblSaldo.Caption, "I", pCharRelleno, 15) '
'strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
'strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
'strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
'strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
'strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
'strLinea(7) = "Operacion/Concepto..: " & "Op.:" & txtOperacion.Text & " Cpt.:" & txtCodigo.Text
'
'If cboDiferenciaApl.Enabled Then
'    strLinea(8) = "Aplica Diferencia ..: " & cboDiferenciaApl.Text
'Else
'    strLinea(8) = "Descripción       ..: " & lblDescripcion.Caption
'
'End If
'
'
'
'strLinea(9) = ""
'strLinea(10) = "Num. Documento    ..:" & lblDocumento.Caption
'strLinea(11) = ""
'
'strSQL = "exec spCxC_OperacionFechaProxPago " & txtOperacion.Text
'Call OpenRecordSet(rsTmp, strSQL, 0)
'  If Not IsNull(rsTmp!fecha_corte) Then
'       strLinea(9) = "Prox.Pago..:" & Format(rsTmp!fecha_corte, "dd/mm/yyyy") & " Cta.(" & rsTmp!Linea & ") " & Format(rsTmp!Monto, "Standard")
'  Else
'       strLinea(9) = "Prox.Pago..: >> <<"
'  End If
'  strLinea(10) = "Notas: " & rsTmp!Notas & ""
'rsTmp.Close
'
'strLinea(10) = Mid(strLinea(10), 1, 80)
'
'
'If dtpFechaCancelacion.Enabled Then
'   strLinea(11) = "Fecha Real Abono  ..: " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
'End If
'
''Registro del Comprobante
'strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
'         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
'         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
'         & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
'         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
'         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
'         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
'         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
'         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
'         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
' Call ConectionExecute(strSQL)
'
' 'ASIENTO
' If curIntC > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
' If curIntM > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
' If curCargo > 0 Then
' 'Detallar Cargos
'   strSQL = "exec spCxC_DocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
'   Call OpenRecordSet(rsTmp, strSQL, 0)
'   Do While Not rsTmp.EOF
'         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Monto * pTipoCambio & ",'C','" & rs!cod_Divisa _
'                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
'                & "','" & rsTmp!Operacion & "','" & rsTmp!cod_Concepto & "','" & vAseDocDeposito & "'"
'         Call ConectionExecute(strSQL)
'         rsTmp.MoveNext
'   Loop
'   rsTmp.Close
' End If
'
'
' If curAmortiza > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
'
'  If curIntC + curIntM + curCargo + curAmortiza > 0 Then
'     'Procesa Formas de Pago (Registro Final / Asiento de Pago)
'      strSQL = "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
'              & "','" & ModuloCajas.mUsuario & "','" & pTipoDoc & "','" & pNumDoc & "','" & ModuloCajas.mUnidad _
'              & "','" & rs!Operacion & "','" & rs!cod_Concepto & "'"
'      Call ConectionExecute(strSQL)
' End If
'
'rs.Close


End Sub



Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String, i As Integer
Dim pOperacion As Long, pFactura As String, pAbono As Currency


Me.MousePointer = vbHourglass

On Error GoTo vError

vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)


'Procesa las Facturas Canceladas
strSQL = ""
With lsw.ListItems
  For i = 1 To .Count
    If .Item(i).Checked Then
        pOperacion = .Item(i)
        pFactura = .Item(i).SubItems(1)
        pAbono = CCur(.Item(i).SubItems(2))
         
        strSQL = strSQL & Space(10) & "exec spCxC_Operacion_Factura_Cancela " & pOperacion & ",'" & pFactura & "'," & pAbono _
               & ",'" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario & "'"
    End If
  Next i
End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
End If


'Procesa Abono + Documento + Asiento
strSQL = "exec spCxC_Operacion_Factura_Cancela_Abono '" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mCaja _
       & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Registra Cancelación de Facturas> Cliente Id: " & feClienteId.Text)

'Imprime el Comprobante
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault

strSQL = " - Abono aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
       & " - Desea Realizar Otra Transacción?"

i = MsgBox(strSQL, vbYesNo)
If i = vbYes Then
    Call sbConsultaCliente
    txtTotalCajas.Text = 0
Else
    Unload Me
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnCajas_Click(Index As Integer)
Dim iRespuesta As Integer

On Error GoTo vError

Select Case Index
  Case 2 'Cancelar
     Call sbConsultaCliente
     
  Case 0 'Desgloce
        If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este Concepto de Cuentas por Cobrar", vbExclamation
           Exit Sub
        End If
                
        If cboDivisas.ListCount = 0 Then
            ModuloCajas.mDivisa = "COL"
        Else
            ModuloCajas.mDivisa = cboDivisas.Text
        End If
        ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        
        If txtTotalCajas.Text <> txtTotalPagar.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

        txtDiferencia.Text = Format((CCur(txtTotalCajas.Text) - CCur(txtTotalPagar.Text)), "Standard")
  
  Case 1 'Aplicar

'    If Not fxVerifica Then Exit Sub
     If CCur(txtTotalCajas.Text) <> CCur(txtTotalPagar.Text) Then
           MsgBox "No se ha recaudado el total a cancelar/pagar de las facturas seleccionadas!", vbInformation
           Exit Sub
     End If
     
     If CCur(txtTotalPagar.Text) = 0 Then
           MsgBox "No se ha indicado ninguna factura a cancelar!", vbInformation
           Exit Sub
     End If
     
     iRespuesta = MsgBox("Esta seguro de realizar el abono a las facturas?", vbYesNo)
     If iRespuesta = vbYes Then
        Call sbAbono
'        Call sbConsultaCliente
     End If


End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
