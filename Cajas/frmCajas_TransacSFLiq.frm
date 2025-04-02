VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCajas_TransacSFLiq 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidación de Saldos a Favor"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   7080
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmCajas_TransacSFLiq.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnCancelar 
         Height          =   495
         Left            =   8400
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         Picture         =   "frmCajas_TransacSFLiq.frx":0727
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3372
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   9852
      _Version        =   1441793
      _ExtentX        =   17378
      _ExtentY        =   5948
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
      FlatScrollBar   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerCaja 
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6555
      Width           =   10140
      _ExtentX        =   17886
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
            TextSave        =   "10:25:p. m."
            Object.ToolTipText     =   "Fecha/Hora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboTipoSaldo 
      Height          =   312
      Left            =   3360
      TabIndex        =   6
      Top             =   4560
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.ComboBox cboTipoLiquidacion 
      Height          =   312
      Left            =   3360
      TabIndex        =   7
      Top             =   4920
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4080
      TabIndex        =   9
      Top             =   360
      Width           =   5892
      _Version        =   1441793
      _ExtentX        =   10393
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2040
      TabIndex        =   10
      Top             =   360
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   8040
      TabIndex        =   11
      Top             =   4920
      Width           =   1932
      _Version        =   1441793
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSFId 
      Height          =   312
      Left            =   8040
      TabIndex        =   12
      Top             =   4560
      Width           =   1932
      _Version        =   1441793
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo a Favor (Id) ..:"
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
      Height          =   195
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Documento:"
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
      Height          =   252
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   4560
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Liquidación:"
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
      Height          =   252
      Index           =   18
      Left            =   1080
      TabIndex        =   3
      Top             =   4920
      Width           =   2172
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6600
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCajas_TransacSFLiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mToken As String


Private Sub btnAplicar_Click()
If Mid(cboTipoLiquidacion.Text, 1, 1) <> "N" And CCur(txtSFId.Text) > 0 Then
   Call sbLiquidaSF
Else
   MsgBox "No se ha indicado un metodo de liquidación del Saldo a Favor!", vbExclamation
End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub cboTipoSaldo_Click()
If vPaso Or cboTipoSaldo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

cboTipoLiquidacion.Clear
cboTipoLiquidacion.AddItem "NO APLICA"
cboTipoLiquidacion.Text = "NO APLICA"

If cboTipoSaldo.Text <> "TODOS" And cboTipoSaldo.ListCount > 0 Then
  strSQL = "select dbo.fxCajas_SaldoFavorTipoLiquidacion('" & cboTipoSaldo.ItemData(cboTipoSaldo.ListIndex) _
         & "','" & glogon.Usuario & "') as 'TipoLiquidacion'"
  Call OpenRecordSet(rs, strSQL)
  Select Case rs!TipoLiquidacion
     Case 0 'No Aplica
     Case 1 'Fondos
        cboTipoLiquidacion.AddItem "Fondos"
      
     Case 2 'Tesoreria
        cboTipoLiquidacion.AddItem "Tesorería"
        
     Case 3 'Ambos
        cboTipoLiquidacion.AddItem "Fondos"
        cboTipoLiquidacion.AddItem "Tesorería"
        
     Case 4 'Efectivo
        cboTipoLiquidacion.AddItem "Efectivo"
        
     Case 7 'Todas
        cboTipoLiquidacion.AddItem "Efectivo"
        cboTipoLiquidacion.AddItem "Fondos"
        cboTipoLiquidacion.AddItem "Tesorería"

     
  End Select
  rs.Close
End If

'Consulta Lista de Casos Disponibles SF para el Usuario actual
Call sbLswSaldosFavor


End Sub

Private Sub Form_Activate()
vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 5

mToken = CStr(Hour(Time))

txtCedula.Text = ModuloCajas.mClienteId
txtNombre.Text = ModuloCajas.mCliente

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

strSQL = "select  rtrim(T.DOC_TIPO) as 'Idx' , rtrim(T.DESCRIPCION) as 'itmX'" _
       & "  from CAJAS_SALDO_FAVOR C inner join CAJAS_SALDOS_FAVOR_TIPOS T on C.DOC_TIPO = T.DOC_TIPO" _
       & "  where C.SALDO > 0 and C.CEDULA = '" & txtCedula.Text & "'" _
       & "  group by T.DOC_TIPO, T.DESCRIPCION" _
       & "  ORDER BY T.DOC_TIPO"
Call sbCbo_Llena_New(cboTipoSaldo, strSQL, False, True)

vPaso = False

With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 1100
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "No.Documento", 2000
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "T.C.", 1000, vbRightJustify
End With


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass


ModuloCajas.mTiquete = ""
txtCedula.Text = ModuloCajas.mClienteId
txtNombre = ModuloCajas.mCliente


txtMonto = 0


Call cboTipoSaldo_Click

MousePointer = vbDefault

End Sub


Private Sub sbLswSaldosFavor()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True
lsw.ListItems.Clear
txtSFId.Text = 0
txtMonto.Text = 0

strSQL = "exec spCajas_SF_Liquidables '" & txtCedula.Text & "','" & cboTipoSaldo.ItemData(cboTipoSaldo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , CStr(rs!Linea))
      itmX.SubItems(1) = rs!DOC_TIPO
      itmX.SubItems(2) = rs!Doc_Numero
      itmX.SubItems(3) = Format(rs!Monto, "Standard")
      itmX.SubItems(4) = Format(rs!Saldo, "Standard")
      itmX.SubItems(5) = rs!cod_Divisa & ""
      itmX.SubItems(6) = rs!TIPO_CAMBIO & ""
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbLiquidaSF()
Dim strSQL  As String, rs As New ADODB.Recordset
Dim i As Long, vIdSaldoFavor As Long
Dim vMetodo As String, vMonto As Currency



On Error GoTo vError

vMetodo = Mid(cboTipoLiquidacion.Text, 1, 1)
vIdSaldoFavor = txtSFId.Text

If vMetodo = "E" Then
  strSQL = "select count(*) as 'Existe'" _
         & " from CAJAS_DEFINICION" _
         & " WHERE isnull(PERMITE_RC,0) = 1 and cod_CAJA = '" & ModuloCajas.mCaja & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
      MsgBox "Esta Caja no permite Retiros en Efectivo!", vbExclamation
      Exit Sub
  End If
  
End If


Me.MousePointer = vbHourglass



Select Case vMetodo
 Case "T" 'Tesoreria
     strSQL = "exec spCajas_SaldoFavorLiquidacionTesoreria " & vIdSaldoFavor & ",'" & glogon.Usuario _
            & "',0,'TE','','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura
     Call ConectionExecute(strSQL)
 
 Case "F" 'Fondos
     strSQL = "exec spCajas_SaldoFavorLiquidacionFondos " & vIdSaldoFavor & ",'" & glogon.Usuario & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura
     Call ConectionExecute(strSQL)
    
 Case "E" 'Retiro de Efectivo en Cajas
     strSQL = "exec spCajas_SaldoFavorLiquidacionRC_Efectivo " & vIdSaldoFavor & ",'" & glogon.Usuario & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura
     Call OpenRecordSet(rs, strSQL)
     Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc)
     
End Select




Call Bitacora("Aplica", "Liquidación de Saldo a Favor: " & vMetodo & " (id." & vIdSaldoFavor & ")")

Me.MousePointer = vbDefault

MsgBox "Saldos a Favor liquidados Satisfactoriamente..!", vbInformation

'Refresca la Lista
Call cboTipoSaldo_Click

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tblAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Aplicar" Then
     If Mid(cboTipoLiquidacion.Text, 1, 1) <> "N" And CCur(txtSFId.Text) > 0 Then
        Call sbLiquidaSF
     Else
        MsgBox "No se ha indicado un metodo de liquidación del Saldo a Favor!", vbExclamation
     End If
Else
    Unload Me
End If

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curMonto As Currency

If vPaso Then Exit Sub

     
 curMonto = CCur(Item.SubItems(4))
 txtMonto.Text = Format(curMonto, "Standard")
 
 txtSFId.Text = Item.Text
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


Me.Caption = "Liquidación de Saldos a Favor ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

StatusBarX.Panels(1).Text = ModuloCajas.mDescripcion
StatusBarX.Panels(2).Text = ModuloCajas.mOficinaDesc

'Inicializa datos Principales
Call sbInicializa

End Sub
