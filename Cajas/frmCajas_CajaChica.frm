VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_CajaChica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retiros en Caja Chica"
   ClientHeight    =   6480
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5052
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   8911
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
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerCaja 
      Interval        =   10
      Left            =   9840
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   6228
      Width           =   11484
      _ExtentX        =   20267
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Caja"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            TextSave        =   "04:49"
            Object.ToolTipText     =   "Fecha/Hora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   8880
      TabIndex        =   1
      Top             =   120
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   4692
      _Version        =   1572864
      _ExtentX        =   8276
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtServicioBusqueda 
      Height          =   312
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   3492
      _Version        =   1572864
      _ExtentX        =   6159
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbCajas 
      Height          =   975
      Left            =   4920
      TabIndex        =   7
      Top             =   5040
      Width           =   6495
      _Version        =   1572864
      _ExtentX        =   11456
      _ExtentY        =   1720
      _StockProps     =   79
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCaja_Aplicar 
         Height          =   615
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Appearance      =   16
         Picture         =   "frmCajas_CajaChica.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnCaja_Cerrar 
         Height          =   615
         Left            =   5640
         TabIndex        =   9
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1291
         _ExtentY        =   1080
         _StockProps     =   79
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
         Transparent     =   -1  'True
         Appearance      =   16
         Picture         =   "frmCajas_CajaChica.frx":07D8
      End
   End
   Begin XtremeSuiteControls.ComboBox cboDivisaActual 
      Height          =   312
      Left            =   6240
      TabIndex        =   10
      Top             =   2160
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   6240
      TabIndex        =   11
      Top             =   2880
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNRef 
      Height          =   312
      Left            =   6240
      TabIndex        =   12
      Top             =   2520
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtServicioCod 
      Height          =   312
      Left            =   6240
      TabIndex        =   13
      Top             =   1440
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecaudadorCod 
      Height          =   312
      Left            =   6240
      TabIndex        =   14
      Top             =   1080
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecaudadorDesc 
      Height          =   312
      Left            =   7560
      TabIndex        =   15
      Top             =   1080
      Width           =   3372
      _Version        =   1572864
      _ExtentX        =   5948
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
   Begin XtremeSuiteControls.FlatEdit txtServicioDesc 
      Height          =   312
      Left            =   7560
      TabIndex        =   16
      Top             =   1440
      Width           =   3372
      _Version        =   1572864
      _ExtentX        =   5948
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
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   1152
      Left            =   6240
      TabIndex        =   17
      Top             =   3240
      Width           =   4812
      _Version        =   1572864
      _ExtentX        =   8488
      _ExtentY        =   2032
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
      BackColor       =   16777215
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboDocumento 
      Height          =   312
      Left            =   6240
      TabIndex        =   22
      Top             =   4560
      Width           =   4812
      _Version        =   1572864
      _ExtentX        =   8493
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto ..:"
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
      Height          =   192
      Index           =   5
      Left            =   4920
      TabIndex        =   25
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa ..:"
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
      Height          =   192
      Index           =   2
      Left            =   4920
      TabIndex        =   24
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FF0000&
      Height          =   192
      Index           =   6
      Left            =   4920
      TabIndex        =   23
      Top             =   4560
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Servicio"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recaudador"
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
      Index           =   21
      Left            =   4920
      TabIndex        =   20
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
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
      Height          =   192
      Index           =   1
      Left            =   4920
      TabIndex        =   19
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Ref ..:"
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
      Height          =   192
      Index           =   4
      Left            =   4920
      TabIndex        =   18
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente .:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servicios ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   216
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   840
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCajas_CajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mToken As String

Private Sub btnCaja_Aplicar_Click()
On Error GoTo vError


If fxValida Then
    Call sbAplicaRetiro
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCaja_Cerrar_Click()
On Error GoTo vError

Call sbInicializa

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cboDivisaActual_Click()
If vPaso Then Exit Sub

Call sbTipoCambioDivisa(cboDivisaActual.ItemData(cboDivisaActual.ListIndex))
End Sub

Private Sub Form_Activate()
 vModulo = 5

End Sub

Private Sub Form_Load()
vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mToken = CStr(Hour(Time))

With lsw.ColumnHeaders
    .Clear
    .Add , , "Concepto", 3500
    .Add , , "Recaudador", 2800
End With


txtCedula.Text = ModuloCajas.mClienteId
txtNombre.Text = ModuloCajas.mCliente

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbAplicaRetiro()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String, i As Integer
Dim vRef_01 As String, vRef_02 As String, vRef_03 As String
Dim strLinea(4), curMonto As Currency, strDivisa As String
Dim pTipoCambio As Currency


On Error GoTo vError

curMonto = CCur(txtMonto.Text)
strDivisa = cboDivisaActual.ItemData(cboDivisaActual.ListIndex)

pTipoCambio = fxCajasTipoCambio(strDivisa)


strSQL = "exec spCajas_ServiciosDatos '" & txtRecaudadorCod.Text & "','" & txtServicioCod.Text _
       & "'," & curMonto & ", '" & ModuloCajas.mCaja & "'"
Call OpenRecordSet(rs, strSQL)



vTipoDoc = cboDocumento.ItemData(cboDocumento.ListIndex)
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)


strLinea(1) = Mid(txtRecaudadorCod.Text & " - " & txtRecaudadorDesc.Text, 1, 80)
strLinea(2) = Mid("N.Ref        ..: " & txtNRef.Text, 1, 80)
strLinea(3) = "Divisa       ..: " & strDivisa
strLinea(4) = Mid("Concepto/Serv..: " & txtServicioCod.Text & " - " & txtServicioDesc.Text, 1, 80)

vRef_01 = txtRecaudadorCod.Text
vRef_02 = txtServicioCod.Text
vRef_03 = Mid(txtNRef.Text, 1, 30)


txtNombre.Text = Mid(txtNombre.Text, 1, 60)

'Control de Documentos v2
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03, cod_oficina" _
        & ",linea1,linea2,linea3,linea4,detalle,documento,cod_caja,cod_apertura)" _
        & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtCedula.Text _
        & "','" & Trim(txtNombre.Text) & "','" & Trim(rs!cod_Concepto) & "'," & curMonto & ",'P','" & vRef_01 _
        & "','" & vRef_02 & "','" & vRef_03 & "','" & ModuloCajas.mOficina & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & txtDetalle.Text & "','" & vAseDocDeposito & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ")"
Call ConectionExecute(strSQL)


With ModuloCajas
    strSQL = "insert CAJAS_SERVICIOS_TRANSAC(Linea, Cod_Caja,Cod_Apertura,Cod_Recaudador,Cod_Servicio,Tipo_Documento,Cod_Transaccion" _
           & ",num_referencia,monto,comision,impuesto,neto, cod_divisa, Tipo_Cambio)" _
           & " Values ( (select isnull(max(Linea),0) + 1 from CAJAS_SERVICIOS_TRANSAC ) " _
           & ",'" & .mCaja & "'," & .mApertura & ",'" & txtRecaudadorCod.Text & "','" & txtServicioCod.Text & "','" & vTipoDoc & "','" & vNumDoc _
           & "','" & Mid(txtNRef.Text, 1, 30) & "'," & rs!Mnt_Bruto & "," & rs!Comision & "," & rs!Impuesto & "," _
           & rs!Mnt_Neto & ",'" & strDivisa & "'," & pTipoCambio & ")"

    Call ConectionExecute(strSQL)
End With

'Asiento:
    'Debita el Concepto de Salida
    strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curMonto * fxSys_Tipo_Cambio_Apl(pTipoCambio) _
            & ",'D','" & strDivisa & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" _
            & rs!Cod_Centro_Costo & "','" & rs!cod_cuenta & "','" & vRef_01 & "','" & vRef_02 & "','" & vRef_03 & "'"
    Call ConectionExecute(strSQL)
    
    'Registra Segun Forma de Pago (Efectivo)
    strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curMonto * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
            & ",'C','" & strDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" _
            & rs!Cod_Centro_Costo & "','" & rs!EF_CTA & "','" & vRef_01 & "','" & vRef_02 & "','" & vRef_03 & "'"
    Call ConectionExecute(strSQL)


'Registra Forma de Pago:

strSQL = "exec spCajas_IntercambioRegistra '" & vTipoDoc & "','" & vNumDoc & "','" & rs!EF_CODIGO & "'," & CCur(txtMonto.Text) _
       & ",'" & rs!EF_CTA & "','" & rs!Cod_Unidad & "','" & glogon.Usuario & "','Retiro en Cajas'"
Call ConectionExecute(strSQL)


Call sbImprimeRecibo(vNumDoc, vTipoDoc)
 
 strSQL = " - Retiro aplicado, con : " & cboDocumento.Text & " ...No.: " & vNumDoc & vbCrLf _
        & " - Desea Realizar Otra Transacción ?"
 
 i = MsgBox(strSQL, vbYesNo)
 If i = vbYes Then
     
        ModuloCajas.mTiquete = ""
        txtCedula.Text = ModuloCajas.mClienteId
        txtNombre = ModuloCajas.mCliente
        
        txtMonto = 0
        txtNRef.Text = ""
        txtDetalle.Text = ""

 Else
     Unload Me
 End If
 

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""

If Not IsNumeric(txtMonto.Text) Then
   txtMonto.Text = "0"
End If

If CCur(txtMonto.Text) <= 0 Then
   vMensaje = vMensaje & "- El Monto de la Transacción no es valido!" & vbCrLf

End If


If fxCajasAperturaEstado = "C" Then
   vMensaje = vMensaje & "- La apertura ..:" & ModuloCajas.mApertura & " de esta caja ha sido cerrada!" & vbCrLf
End If

Call sbSIFCleanTxtInject(txtDetalle)


If Len(vMensaje) = 0 Then
    fxValida = True
Else
    fxValida = False
    MsgBox vMensaje, vbExclamation
End If

End Function

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

  
  txtServicioCod.Text = Item.Tag
  txtServicioDesc.Text = Item.Text
  
  txtRecaudadorCod.Text = Item.ListSubItems.Item(1).Tag
  txtRecaudadorDesc.Text = Item.SubItems(1)


txtCedula.SetFocus

End Sub




Private Sub sbBuscaServicios()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select S.cod_servicio,S.descripcion as 'ServicioDesc',R.Cod_recaudador,R.descripcion as 'RecaudadorDesc'" _
        & " from cajas_servicios_asignados X " _
        & " inner join cajas_servicios S on X.cod_Recaudador = S.Cod_Recaudador and X.cod_Servicio = S.cod_Servicio" _
        & " inner join cajas_recaudador R on S.cod_recaudador = R.cod_recaudador " _
        & " where X.cod_Caja = '" & ModuloCajas.mCaja & "' and S.descripcion like '%" & Trim(txtServicioBusqueda.Text) _
        & "%' and S.cod_Concepto  in('CAJ002')"
Call OpenRecordSet(rs, strSQL)

vPaso = True

lsw.ListItems.Clear
Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!ServicioDesc)
        itmX.SubItems(1) = rs!RecaudadorDesc
        itmX.Tag = rs!COD_SERVICIO
        itmX.ListSubItems(1).Tag = rs!COD_RECAUDADOR
 rs.MoveNext
Loop
rs.Close

txtServicioCod.Text = ""
txtServicioDesc.Text = ""
txtRecaudadorCod.Text = ""
txtRecaudadorDesc.Text = ""

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub TimerCaja_Timer()

TimerCaja.Interval = 0
TimerCaja.Enabled = False

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

'Paso 3: Continuar con Barra de Información
'lblInfoApertura.Caption = ModuloCajas.mApertura
'lblInfoCaja.Caption = ModuloCajas.mCaja
'lblInfoUsuario.Caption = ModuloCajas.mUsuario

Me.Caption = "Retiros en Caja Chica  ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario


StatusBarX.Panels(1).Text = ModuloCajas.mDescripcion
StatusBarX.Panels(2).Text = ModuloCajas.mOficinaDesc

'Inicializa datos Principales
Call sbInicializa

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "select Cedula,CedulaR,nombre from socios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    gBusquedas.Convertir = "N"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtCedula.Text = gBusquedas.Resultado
        txtNombre.Text = gBusquedas.Resultado3
    
        ModuloCajas.mClienteId = txtCedula.Text
        ModuloCajas.mCliente = txtNombre.Text
    End If
End If

End Sub

Private Sub txtCedula_LostFocus()
 txtNombre.Text = fxNombre(txtCedula.Text)
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
 txtMonto.SelStart = Len(txtMonto)
Exit Sub

vError:
 txtMonto = 0
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDetalle.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
txtMonto = Format(txtMonto, "Standard")
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

End Sub


Private Sub sbInicializa()
Dim strSQL As String

ModuloCajas.mTiquete = ""
txtCedula.Text = ModuloCajas.mClienteId
txtNombre.Text = ModuloCajas.mCliente


txtMonto.Text = 0
txtNRef.Text = ""
txtDetalle.Text = ""

Call sbCargaDivisa(cboDivisaActual)

Call sbBuscaServicios

strSQL = "select rtrim(Doc.Tipo_Documento) as 'IdX', rtrim(Doc.Descripcion) as 'Itmx'" _
       & " from Cajas_Documentos Cj inner join SIF_Documentos Doc on Cj.Tipo_Documento = Doc.Tipo_Documento" _
       & " Where Cj.Cod_Caja = '" & ModuloCajas.mCaja & "' and Doc.Tipo_Documento in('CAJA','CAJRE')"
Call sbCbo_Llena_New(cboDocumento, strSQL, False, False)

End Sub


Private Sub sbCargaDivisa(cbo As Object)
Dim strSQL As String, rs As New ADODB.Recordset

vPaso = True
strSQL = "Select rtrim(Cod_Divisa) as 'IdX' , rtrim(descripcion) as 'itmx' from cntx_divisas where cod_contabilidad = " & GLOBALES.gEnlace & " order by cod_divisa"
    Call sbCbo_Llena_New(cbo, strSQL, False, False)
vPaso = False

Call cboDivisaActual_Click

End Sub


Private Sub sbTipoCambioDivisa(vCodigo As String)
'Dim strSQL As String, rs As New ADODB.Recordset
'
'strSQL = "select TC_VENTA,TC_COMPRA from cntx_divisas where cod_divisa = '" & vCodigo & "' and cod_contabilidad= " & GLOBALES.gEnlace & ""
'Call OpenRecordSet(rs, strSQL)
'
'If Not rs.EOF Then
'  StatusBar.Panels(2).Text = Format(rs!tc_venta, "Standard")
'  StatusBar.Panels(4).Text = Format(rs!tc_compra, "Standard")
'Else
'  StatusBar.Panels(2).Text = 0
'  StatusBar.Panels(4).Text = 0
'End If
'rs.Close

End Sub


Private Sub txtServicioBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)
  Call sbBuscaServicios
End Sub


