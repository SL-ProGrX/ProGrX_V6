VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_ConsultaCrdExc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crédito con Garantía en los Excedentes"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
      _ExtentY        =   7429
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
      HideSelection   =   0   'False
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   2187
            MinWidth        =   2187
            TextSave        =   "14/12/2023"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6615
            MinWidth        =   6615
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton cmdFormalizar 
      Height          =   612
      Left            =   8880
      TabIndex        =   11
      Top             =   4560
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Formalizar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_ConsultaCrdExc.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   6480
      TabIndex        =   15
      Top             =   2760
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboCuenta 
      Height          =   312
      Left            =   6480
      TabIndex        =   16
      Top             =   3120
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
      Height          =   312
      Left            =   8280
      TabIndex        =   13
      Top             =   2400
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4260
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
   Begin XtremeSuiteControls.ComboBox cboRecursos 
      Height          =   312
      Left            =   6480
      TabIndex        =   14
      Top             =   3600
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   8280
      TabIndex        =   17
      Top             =   1080
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMontoAplicar 
      Height          =   312
      Left            =   8280
      TabIndex        =   18
      Top             =   1440
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMontoGirar 
      Height          =   312
      Left            =   8280
      TabIndex        =   19
      Top             =   1800
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDisponibleRecursos 
      Height          =   312
      Left            =   6480
      TabIndex        =   20
      Top             =   3960
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
   Begin VB.Label lblMora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   5160
      TabIndex        =   10
      Top             =   4560
      Width           =   3612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Girar...: "
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
      Index           =   4
      Left            =   6600
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible"
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
      Index           =   3
      Left            =   5280
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recurso"
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
      Index           =   2
      Left            =   5280
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   10572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emitir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   13
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Index           =   14
      Left            =   5280
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
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
      Index           =   15
      Left            =   5280
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a aplicar  ...: "
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto disponible ...: "
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
      Left            =   6600
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_ConsultaCrdExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCedula As String, vPaso As Boolean, mLinea As String, mGarantia As String
Dim mTasaDiaria As Double, mDias As Long, mPoliza As Double, mMora As Currency
Dim strSQL As String, rs As New ADODB.Recordset


Private Function fxLineaExcedenteCodigo() As String
Dim pResultado As String

pResultado = ""


strSQL = "select rtrim(VALOR) as 'Valor'  From EXC_PARAMETROS" _
       & " Where COD_PARAMETRO = '05'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    pResultado = rs!Valor
End If

fxLineaExcedenteCodigo = pResultado
End Function


Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 And cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & mCedula & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  cboTipoDocumento.SetFocus
End If
End Sub


Private Sub cboRecursos_Click()

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spCRDDisponibleRecurso '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "','" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtDisponibleRecursos = Format(rs!Disponible, "Standard")
Else
    txtDisponibleRecursos = 0
End If
rs.Close

Me.MousePointer = vbDefault


End Sub

Private Sub cmdFormalizar_Click()
Dim i As Integer

 If fxVerificaFormalizacion Then
    
     i = MsgBox("Esta seguro que desea >> formalizar << esta Operación", vbYesNo)
     If i = vbYes Then
         Call sbFormalizar
     End If
     
End If 'Verificacion de Formalizacion

End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()
vModulo = 3

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mCedula = GLOBALES.gTag

mLinea = fxLineaExcedenteCodigo

stBar.Panels(3).Text = GLOBALES.gOficina

Call sbInicializa
Call sbConsultaInicial

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbInicializa()


Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select isnull(sum(intc+intm+cargos+amortiza),0) as 'Mora', dbo.fxCrdPersonaValidaCondiciones('" & mCedula & "') as 'Condiciones'" _
       & " From VISTA_MOROSIDAD" _
       & " where cedula = '" & mCedula & "'"
Call OpenRecordSet(rs, strSQL)
   Select Case rs!Condiciones
      Case 0 'Pasa la validacion
        lblMora.Visible = False
        mMora = 0
      Case 1 'Presenta Mora
        mMora = rs!Mora
        If mMora > 0 Then
         lblMora.Visible = True
         lblMora.Caption = "Esta persona tienen una morosidad de : " & Format(rs!Mora, "Standard") & vbCrLf _
                          & " -> No puede formalizar desde esta opción!"
        End If
      Case 2 'Traslado de Deudas
        lblMora.Visible = True
        lblMora.Caption = "Esta persona Presenta Operaciones con Traslado de Deudas" & vbCrLf _
                         & " -> No puede formalizar desde esta opción!"
      Case 3 'Cobro Judicial
        lblMora.Visible = True
        lblMora.Caption = "Esta persona Tiene Operaciones en Cobro Judicial!" & vbCrLf _
                         & " -> No puede formalizar desde esta opción!"
      Case 4 'Caso Bloqueado
        lblMora.Visible = True
        lblMora.Caption = "Esta persona se encuentra bloqueada para nuevas operaciones " & vbCrLf _
                         & " -> No puede formalizar desde esta opción!"
   
   End Select
rs.Close



strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


strSQL = " select G.cod_grupo as 'IdX',rtrim(G.descripcion) as 'ItmX'" _
       & " from catalogo_grupos G inner join catalogo_asignaGrp A on G.cod_grupo = A.cod_grupo" _
       & " where G.estado = 1 and A.codigo = '" & mLinea & "'"
Call sbCbo_Llena_New(cboRecursos, strSQL, False, False)


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("RC")

cboTipoDocumento.Text = fxTipoDocumento("TE")

vPaso = False


Call cboRecursos_Click
Call cboBanco_Click

'txtMontoAplicar.SetFocus

MousePointer = vbDefault

End Sub


Private Sub sbConsultaInicial()

Dim i As Integer
Dim itmX As ListViewItem, frmX As Form

Me.MousePointer = vbHourglass


lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "", 2100
    .Add , , "", 1040, vbCenter
    .Add , , "", 1500, vbRightJustify
End With
lsw.HideSelection = True

strSQL = "exec spVoxExcedenteCredito '" & mCedula & "'"
Call OpenRecordSet(rs, strSQL)


stBar.Panels(2).Text = UCase("[PERIODO: " & Format(rs!periodo_de, "YYYY-MM") & " : " & Format(rs!periodo_hasta, "YYYY-MM") & " ][MES APL: " & fxConvierteMES(rs!mes_aplicado) & "]")

Set itmX = lsw.ListItems.Add(, , "Excedente Bruto")
    itmX.SubItems(1) = "SUM"
    itmX.SubItems(2) = Format(rs!bruto, "Standard")
    itmX.ForeColor = vbBlue
    
Set itmX = lsw.ListItems.Add(, , "(-) Capit.General")
    itmX.SubItems(1) = rs!porCapGen & " %"
    itmX.SubItems(2) = Format(rs!Capitalizacion, "Standard")

Set itmX = lsw.ListItems.Add(, , "(-) Impuesto Renta")
    itmX.SubItems(1) = rs!porRenta & " %"
    itmX.SubItems(2) = Format(rs!Renta, "Standard")

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = "__________"

Set itmX = lsw.ListItems.Add(, , "Base Crédito")
    itmX.SubItems(1) = rs!PorAcumulado & " %"
    itmX.SubItems(2) = Format(rs!Base, "Standard")
    itmX.ForeColor = vbBlue

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = "=========="


Set itmX = lsw.ListItems.Add(, , "(-) Saldos")
    itmX.SubItems(1) = "SUM"
    itmX.SubItems(2) = Format(rs!Saldos, "Standard")


Set itmX = lsw.ListItems.Add(, , "(-) Capit.Individual")
    itmX.SubItems(1) = rs!porCapInd & " %"
    itmX.SubItems(2) = Format(rs!capIndividual, "Standard")

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = "__________"

'Set itmx = lsw.ListItems.Add(, , "Disponible Bruto")
'    itmx.SubItems(1) = "SUM"
'    itmx.SubItems(2) = Format(rs!neto + rs!disponibleBruto, "Standard")
    
Set itmX = lsw.ListItems.Add(, , "Disponible")
    itmX.SubItems(1) = "SUM"
    itmX.SubItems(2) = Format(rs!Neto, "Standard")
    itmX.ForeColor = vbBlue
    
Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = "=========="
    
Set itmX = lsw.ListItems.Add(, , "(-) Intereses (" & rs!Dias & " dias)")
    itmX.SubItems(1) = rs!Tasa & " %"
    itmX.SubItems(2) = Format(rs!Intereses, "Standard")

Set itmX = lsw.ListItems.Add(, , "Giro Máximo Neto")
    itmX.SubItems(1) = "SUM"
    itmX.SubItems(2) = Format(rs!Giro_Maximo, "Standard")
    itmX.ForeColor = vbBlue
    
lblCliente.Caption = Trim(mCedula) & " - " & Trim(rs!Nombre & "")

mTasaDiaria = rs!Tasa / 36000
mDias = rs!Dias
mPoliza = rs!PolizaFactor

MousePointer = vbDefault

'Enviar el Monto Base para Creditos, no el neto ya que se le deben de aplicar los saldos en refundiciones
txtMonto.Text = Format(rs!Neto, "Standard")
txtMontoAplicar.Text = Format(rs!Neto, "Standard")
txtMontoGirar.Text = Format(rs!Neto - (rs!Neto * mTasaDiaria * mDias), "Standard")

rs.Close

End Sub




Private Function fxVerificaFormalizacion() As Boolean
Dim rsX As New ADODB.Recordset
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim curDisponible As Currency, curGiros As Currency
Dim curMontoTmp As Currency, vPriDeducCorte As Long
Dim vMensaje As String

vMensaje = ""

fxVerificaFormalizacion = True

vFecha = fxFechaServidor

'strSQL = "select MAX(proceso) as 'Proceso' From PRM_BITACORA" _
'       & " where COD_INSTITUCION in(select COD_INSTITUCION  from SOCIOS where CEDULA = '" & mCedula _
'       & "') and GESTION = 'E' and TRANSACCION = '02'"
'Call OpenRecordSet(rsX, strSQL, 0)
'If IsNull(rsX!Proceso) Then
'   vPriDeducCorte = GLOBALES.glngFechaCR
'Else
'   vPriDeducCorte = rsX!Proceso
'End If
'
'rsX.Close

If lblMora.Visible Then
    vMensaje = vMensaje & vbCrLf & "- Esta persona tiene Operaciones atrasadas/Traslados/Cobro Judicial -> no puede formalizarse desde acá..."
  
End If

'Verifica que si la salida es por transferencia, la cuenta de ahorros no este en blanco
If fxTipoDocumento(cboTipoDocumento.Text) = "TE" Then
  If cboCuenta.ListCount = 0 Or cboCuenta.Text = "" Then
    vMensaje = vMensaje & vbCrLf & "- No se ha especificado una cuenta de ahorros para realizarle el depósito..."
  End If
End If



'Revision de Garantia en Excedentes
If Not IsNumeric(txtMontoAplicar.Text) Then
         vMensaje = vMensaje & vbCrLf & "- El monto solicitado no es válido..."
Else
    If CCur(txtMontoAplicar.Text) > CCur(txtMonto.Text) Or CCur(txtMontoAplicar.Text) < 0 Then
         vMensaje = vMensaje & vbCrLf & "- El monto solicitado sobrepasa el disponible o no es válido..."
    End If

    If CCur(txtDisponibleRecursos.Text) < CCur(txtMontoAplicar.Text) Then
       vMensaje = vMensaje & vbCrLf & " - No Hay disponible en el Recurso, para desembolsar esta Operación..."
       vMensaje = vMensaje & vbCrLf & " - Monto a Girar : " & Format(CCur(txtMontoAplicar.Text), "Standard") & " - Disponible :  " & Format(CCur(txtDisponibleRecursos.Text), "Standard")
       vMensaje = vMensaje & vbCrLf & " - Monto Faltante para Girar: " & Format(CCur(txtMontoAplicar.Text) - CCur(txtDisponibleRecursos.Text), "Standard")
    End If

        'Retiros en Cajas> Validacion
        If fxTipoDocumento(cboTipoDocumento.Text) = "RC" Then
          strSQL = "select Valor from CAJAS_PARAMETROS  where cod_parametro = '15'"
          Call OpenRecordSet(rsX, strSQL)
          If IsNumeric(rsX!Valor) Then
                If rsX!Valor < CCur(txtMontoAplicar.Text) Then
                    vMensaje = vMensaje & vbCrLf & "- El Monto Máximo para Retiros de Efectivos en Cajas es de " _
                           & Format(rsX!Valor, "Standard") & ", Informe a su Administrador!"
                End If
          Else
            vMensaje = vMensaje & vbCrLf & "- No se ha configurado el Monto para Retiros de Efectivos en Cajas, Informe a su Administrador!"
          End If
        End If


End If




If Len(vMensaje) > 0 Then
  fxVerificaFormalizacion = False
  MsgBox vMensaje, vbExclamation
Else
  fxVerificaFormalizacion = True
End If

End Function






Private Sub sbFormalizar()
Dim pOperacion As Long

On Error GoTo vError

Me.MousePointer = vbHourglass



strSQL = "exec spCrdCreditoExcedentesRapido '" & mLinea & "','" & mCedula & "'," & CCur(txtMontoAplicar.Text) _
       & "," & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & fxTipoDocumento(cboTipoDocumento.Text) _
       & "','" & cboCuenta.ItemData(cboCuenta.ListIndex) & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular & "','" & glogon.AppName & "'"
Call OpenRecordSet(rs, strSQL)
 pOperacion = rs!Operacion
rs.Close

'BITACORA
Call Bitacora("Registra", "Formalización de la OP: " & pOperacion)

'Imprime Boleta de Formalizacion
Call sbCrdSGTBoletaFormaliza(pOperacion)

Me.MousePointer = vbDefault

MsgBox "Formalización Aplicada Satisfactoriamente...", vbInformation

Unload Me

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtMontoAplicar_GotFocus()
On Error GoTo vError
  txtMontoAplicar.Text = CCur(txtMontoAplicar.Text)
vError:
End Sub

Private Sub txtMontoAplicar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoDocumento.SetFocus
End Sub

Private Sub txtMontoAplicar_LostFocus()
On Error GoTo vError
  txtMontoAplicar.Text = Format(CCur(txtMontoAplicar.Text), "Standard")
  Call sbCalculaMontos("Ma", CCur(txtMontoAplicar.Text))
vError:
End Sub


Private Sub txtMontoGirar_GotFocus()
On Error GoTo vError
  txtMontoGirar.Text = CCur(txtMontoGirar.Text)
vError:
End Sub

Private Sub txtMontoGirar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoDocumento.SetFocus
End Sub

Private Sub txtMontoGirar_LostFocus()
On Error GoTo vError
  txtMontoGirar.Text = Format(CCur(txtMontoGirar.Text), "Standard")
  Call sbCalculaMontos("Mg", CCur(txtMontoGirar.Text))
vError:
End Sub


Private Sub sbCalculaMontos(pTipo As String, pMonto As Currency)
Dim curIntereses As Currency, curPoliza As Currency
Dim i As Integer, curResultado As Currency, curTemp As Currency

On Error GoTo vError

Select Case pTipo 'Tipo de Monto que se recibe
    Case "Ma" 'Monto a aplicar
        curIntereses = pMonto * mTasaDiaria * mDias
        curPoliza = pMonto * mPoliza
        
        txtMontoGirar.Text = Format(pMonto - Round(curIntereses + curPoliza), "Standard")
        
    Case "Mg" 'Monto a girar
        curIntereses = pMonto * mTasaDiaria * mDias
        curPoliza = pMonto * mPoliza
        curResultado = pMonto + Round(curIntereses, 2) + Round(curPoliza, 2)
      
      For i = 1 To 10
        curIntereses = curResultado * mTasaDiaria * mDias
        curPoliza = curResultado * mPoliza
        
        'Monto a Girar Temporal
        curTemp = curResultado - Round(curIntereses + curPoliza, 2)
        
        'Revisa si el resultado es mayor o menor al esperado y realiza ajuste a la base
        If curTemp > pMonto Then
            curTemp = (curTemp - pMonto) / 2
            curResultado = curResultado - curTemp
        Else
            curTemp = (pMonto - curTemp) / 2
            curResultado = curResultado + curTemp
        
        End If
      Next i
      
        'Actualiza Campos
        txtMontoAplicar.Text = Format(curResultado, "Standard")
        
        curIntereses = curResultado * mTasaDiaria * mDias
        curPoliza = curResultado * mPoliza
        
        txtMontoGirar.Text = Format(curResultado - Round(curIntereses + curPoliza, 2), "Standard")

End Select


vError:
End Sub

