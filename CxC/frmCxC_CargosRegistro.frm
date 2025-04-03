VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_CargosRegistro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Cargos a Cuentas por Cobrar"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8655
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   372
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   312
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7429
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   7560
      TabIndex        =   6
      Top             =   1080
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   1092
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   8412
      _Version        =   1441793
      _ExtentX        =   14838
      _ExtentY        =   1926
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCargoReposicion 
         Height          =   492
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplicar Cargo por Reposición"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   495
         Left            =   6600
         TabIndex        =   11
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Aplicar"
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
         Picture         =   "frmCxC_CargosRegistro.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   792
      Left            =   1440
      TabIndex        =   12
      Top             =   2160
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   1397
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboCargo 
      Height          =   312
      Left            =   1440
      TabIndex        =   13
      Top             =   1680
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   6480
      TabIndex        =   14
      Top             =   3120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto del Cargo a Registrar..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   3
      Left            =   3000
      TabIndex        =   17
      Top             =   3120
      Width           =   2652
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo ..:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   264
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1452
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12492
   End
End
Attribute VB_Name = "frmCxC_CargosRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboCargo_Click()
If vPaso Or cboCargo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset



strSQL = "select COD_CARGO,DESCRIPCION,COD_CUENTA " _
       & " From CxC_CARGOS where cod_cargo = '" & cboCargo.ItemData(cboCargo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

    txtMonto.Text = Format(0, "Standard")
    txtMonto.Tag = rs!cod_cuenta

rs.Close

End Sub

Private Sub chkCargoReposicion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If chkCargoReposicion.Value = vbChecked Then
   cboCargo.Enabled = False
   txtMonto.Locked = True
   txtNotas.Locked = True
   strSQL = "select dbo.fxCxC_CuentaCargoReposicion(" & txtOperacion.Text & ",null) as 'Cargo'"
   Call OpenRecordSet(rs, strSQL)
     txtMonto.Text = Format(rs!Cargo, "Standard")
   rs.Close
Else
   cboCargo.Enabled = True
   txtMonto.Locked = False
   txtNotas.Locked = False
   txtMonto.Text = Format(0, "Standard")

End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

If chkCargoReposicion.Value = vbUnchecked Then
    strSQL = "exec spCxC_CuentaCargoAdd " & txtOperacion.Text & "," & CCur(txtMonto.Text) & ",'" & GLOBALES.gOficinaUnidad _
           & "','" & GLOBALES.gOficinaCentroCosto & "','" & Mid(txtNotas.Text, 1, 59) & "','" & glogon.Usuario _
           & "','','" & cboCargo.ItemData(cboCargo.ListIndex) & "',0"
Else
    strSQL = "exec spCxC_CuentaCargoReposicion " & txtOperacion.Text & ",'" & glogon.Usuario _
           & "','" & GLOBALES.gOficinaUnidad & "','" & GLOBALES.gOficinaCentroCosto & "',Null"
End If
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Cargo registrado satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 31

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
strSQL = "select rtrim(COD_CARGO) as 'IdX', rtrim(DESCRIPCION) as 'Itmx'" _
       & " From CARGOS_ADICIONALES where TIPO = 'M'"
Call sbCbo_Llena_New(cboCargo, strSQL, False, True)
vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()

txtCedula.Text = ""
txtNombre.Text = ""
txtCodigo.Text = ""
txtLineaDesc.Text = ""

txtOpex.Text = ""
txtProceso.Text = ""

txtNotas.Text = ""

Call chkCargoReposicion_Click
End Sub

Private Sub sbCargaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
Me.MousePointer = vbHourglass

Call sbLimpia

strSQL = "select Soc.cedula,Soc.nombre,Cat.cod_concepto,Cat.descripcion,Reg.proceso,Reg.num_documento" _
       & " from CxC_Personas Soc inner join CxC_Cuentas Reg on Soc.cedula = Reg.cedula" _
       & " inner join CxC_Conceptos Cat on Reg.cod_Concepto = Cat.cod_Concepto" _
       & " where Reg.estado = 'A' and Reg.Operacion = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
If rs.BOF And rs.EOF Then
   Me.MousePointer = vbDefault
   MsgBox "No se encontró operación activa...!", vbExclamation
   Exit Sub
Else
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    txtCodigo.Text = rs!cod_Concepto
    txtLineaDesc.Text = rs!Descripcion
    txtOpex.Text = Trim(rs!Num_Documento)
    txtProceso.Text = fxProcesoOperacion(rs!Proceso)
    txtNotas.SetFocus
End If

rs.Close



Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
    txtMonto.Text = CCur(txtMonto.Text)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And cmdAplicar.Enabled Then cmdAplicar.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
    txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtOperacion_Change()
 Call sbLimpia
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbCargaOperacion
End If
End Sub


