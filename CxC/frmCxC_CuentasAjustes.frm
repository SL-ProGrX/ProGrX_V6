VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxC_CuentasAjustes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CxC > Ajustes ...: Plan de Pagos de la Cuenta"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2172
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   3831
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   612
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   5280
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Días de Atraso según Fecha del Documento"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   160
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   160
      _Version        =   1441793
      _ExtentX        =   282
      _ExtentY        =   282
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFechaDocumento 
      Height          =   312
      Left            =   5280
      TabIndex        =   3
      Top             =   6120
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   672
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   7452
      _Version        =   1441793
      _ExtentX        =   13144
      _ExtentY        =   1185
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
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2040
      TabIndex        =   5
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3720
      TabIndex        =   6
      Top             =   840
      Width           =   5772
      _Version        =   1441793
      _ExtentX        =   10181
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9123
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
      Left            =   2040
      TabIndex        =   8
      Top             =   840
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   8880
      TabIndex        =   10
      Top             =   1200
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
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   612
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   5760
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Registro de Cargos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   612
      Index           =   2
      Left            =   600
      TabIndex        =   12
      Top             =   6240
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Dias de Mora"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAjustar 
      Height          =   732
      Left            =   8040
      TabIndex        =   13
      Top             =   5880
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Ajustar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxC_CuentasAjustes.frx":0000
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   21
      Top             =   6936
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Garantía"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Recursos"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   276
      Left            =   600
      TabIndex        =   20
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label lblAjusteDias 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ajustar Plan de Pagos (Dias de Intereses) a la Fecha del Documento "
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
      Height          =   1092
      Left            =   3960
      TabIndex        =   19
      Top             =   5520
      Width           =   3852
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Ajustes al Plan de Pagos ...:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Left            =   600
      TabIndex        =   17
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Left            =   600
      TabIndex        =   16
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
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
      Left            =   600
      TabIndex        =   15
      Top             =   1560
      Width           =   1332
   End
   Begin XtremeShortcutBar.ShortcutCaption lblMoraTexto 
      Height          =   372
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Seleccione las Cuotas Morosas para Anulación y Luego Presione Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCxC_CuentasAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String

On Error Resume Next

strSQL = "select R.*,S.nombre,C.descripcion" _
       & ",Ofi.Descripcion as 'OficinaX',Cnt.Descripcion as 'Contrato',Pag.Nombre as 'Pagador'" _
       & " from CxC_Cuentas R inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto" _
       & " inner join CxC_Personas S on R.cedula = S.cedula " _
       & " left join CxC_Contratos Cnt on R.Cod_Contrato = Cnt.Cod_Contrato" _
       & " left Join CxC_Personas Pag on R.cedula_pagador = Pag.cedula" _
       & " left join SIF_Oficinas Ofi on R.cod_oficina = Ofi.cod_Oficina" _
       & " where R.estado = 'A' and R.proceso <> 'J' and R.Operacion =" & txtOperacion
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    
   
 txtCedula.Text = rs!Cedula
 txtCodigo.Text = rs!cod_Concepto
 txtNombre.Text = rs!Nombre
 txtDescripcion.Text = rs!Descripcion
 txtOpex.Text = rs!Num_Documento
 
 StatusBarX.Panels.Item(1) = rs!OficinaX & ""
 StatusBarX.Panels.Item(2) = rs!Contrato & ""
 StatusBarX.Panels.Item(3) = rs!Pagador & ""
End If

rs.Close

Call OptX_Click(0)

End Sub

Private Sub btnAjuste_Click()
Select Case True
   Case OptX.Item(0).Value 'Ajuste a la fecha del Documento
      Call sbAjustaFechaDocumento
   Case OptX.Item(1).Value 'Registro de Cargos
      If Len(txtNotas.Text) < 15 Then
        MsgBox "No ha especificado una Nota válida para registrar el cambio...?"
        Exit Sub
      End If
      Call sbEliminaCargos
   Case OptX.Item(2).Value 'Dias de Mora
      If Len(txtNotas.Text) < 15 Then
        MsgBox "No ha especificado una Nota válida para registrar el cambio...?"
        Exit Sub
      End If
      Call sbEliminaMora
End Select

If GLOBALES.gTag2 = 1 Then
  Unload Me
End If
End Sub

Private Sub chkTodas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = vbChecked
Next i

End Sub

Private Sub Form_Load()
Dim iDias As Integer, vFechaHoy As Date

vModulo = 31

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vFechaHoy = fxFechaServidor
iDias = fxCrdParametro("32")

txtOperacion.Text = GLOBALES.gTag
GLOBALES.gTag2 = 0

Call sbCargaOperacion


dtpFechaDocumento.Value = vFechaHoy
dtpFechaDocumento.MinDate = DateAdd("d", (iDias * -1), dtpFechaDocumento.Value)
dtpFechaDocumento.MaxDate = dtpFechaDocumento.Value

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbAjustaFechaDocumento()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "exec spCxC_CuentaIntereses " & txtOperacion.Text & ",'','" & Format(dtpFechaDocumento.Value, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)

'Call sbBitacoraCredito("06", "Ajusta según Fecha Documento", "C", txtOperacion _
'                    , txtCodigo.Text, "Fecha de Corte del Documento : " & Format(dtpFechaDocumento.Value, "dd/mm/yyyy"))

GLOBALES.gTag2 = 1

Me.MousePointer = vbDefault

MsgBox "Ajuste de Fecha de Documento en Plan de Pagos Realizado Satisfactoriamente...!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLlenaMorosidad()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError


lblMoraTexto.Caption = ">> Seleccione las Cuotas Morosas para Ajustar <<"

chkTodas.Enabled = True

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear


lsw.ColumnHeaders.Add , , "Linea", 900
lsw.ColumnHeaders.Add , , "Proceso", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 1200
lsw.ColumnHeaders.Add , , "Int.Cor", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Int.Mor.", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Principal", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Cargos", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Dias", 1200, vbCenter

rs.CursorLocation = adUseServer


strSQL = "select * From CxC_Cuentas_Mov" _
       & " where Dias_Mora > 0 AND ESTADO = 'A' AND Operacion = " & txtOperacion.Text

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Linea)
     itmX.SubItems(1) = Format(rs!Fecha_Inicio, "yyyy-mm")
     itmX.SubItems(2) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
     itmX.SubItems(3) = Format(rs!Int_Cor, "Standard")
     itmX.SubItems(4) = Format(rs!Int_Mor, "Standard")
     itmX.SubItems(5) = Format(rs!Principal, "Standard")
     itmX.SubItems(6) = Format(rs!Cargos, "Standard")
     itmX.SubItems(7) = Format(rs!Dias_Mora, "Standard")
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub sbLlenaCargos()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError


lblMoraTexto.Caption = ">> Seleccione los Cargos a Eliminar <<"

chkTodas.Enabled = True

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "ID", 900
lsw.ColumnHeaders.Add , , "Proceso", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 1200
lsw.ColumnHeaders.Add , , "Usuario", 1200
lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Detalle", 3200
lsw.ColumnHeaders.Add , , "ID.Mora", 1000


rs.CursorLocation = adUseServer
strSQL = "select * from CxC_Cuentas_Cargos where Operacion = " & txtOperacion.Text _
       & " and Monto = Saldo"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_Cargo)
     itmX.SubItems(1) = Format(rs!Registro_Fecha, "####-##")
     itmX.SubItems(2) = Format(rs!Registro_Fecha, "dd/mm/yyyy")
     itmX.SubItems(3) = rs!Registro_Usuario
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Notas
     itmX.SubItems(6) = rs!Id_Cargo
     
     itmX.Checked = chkTodas.Value
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub OptX_Click(Index As Integer)

lblAjusteDias.Visible = False
dtpFechaDocumento.Visible = False

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lblMoraTexto.Caption = ""
chkTodas.Enabled = False

Select Case True
   Case OptX.Item(0).Value 'Ajuste a la fecha del Documento
        lblAjusteDias.Visible = True
        dtpFechaDocumento.Visible = True
        
   Case OptX.Item(1).Value 'Registro de Cargos
      Call sbLlenaCargos
   Case OptX.Item(2).Value 'Dias de Mora
      Call sbLlenaMorosidad

End Select

End Sub



Private Sub sbEliminaMora()
Dim strSQL As String, itmX As ListItem, lng As Long

On Error GoTo vError


With lsw.ListItems
    For lng = 1 To .Count
       If .Item(lng).Checked Then
            strSQL = "update CxC_Cuentas_Mov set Dias_Mora = 0, int_Mor = 0 where Linea = " _
                   & .Item(lng).Text & " and Operacion = " & txtOperacion.Text
            Call ConectionExecute(strSQL)
          

          strSQL = "Int.Mor..: " & .Item(lng).SubItems(4) & "   Dias..: " & .Item(lng).SubItems(7) & "    Notas..: " & txtNotas.Text
            
'          Call sbBitacoraCredito("06", "Id..:" & .Item(lng).Text, "C", txtOperacion, txtCodigo.Text, strSQL)
          Call Bitacora("Anula", "Morosidad OP: " & txtOperacion & " Linea:" & .Item(lng).Text)
       
          GLOBALES.gTag2 = 1 'Bandera que indica que se realizó un Ajuste
       
       End If
    Next lng
End With


MsgBox "Reversiones realizadas Satisfactoriamente...", vbInformation


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub sbEliminaCargos()
Dim strSQL As String, itmX As ListItem, lng As Long

On Error GoTo vError

With lsw.ListItems
    For lng = 1 To .Count
       If .Item(lng).Checked Then
                
            strSQL = "exec spCxC_CuentaCargoElimina " & txtOperacion.Text & "," & .Item(lng).Text
            Call ConectionExecute(strSQL)
                
'            Call sbBitacoraCredito("21", .Item(lng).SubItems(5), "C", txtOperacion, txtCodigo.Text, "Monto..: " & .Item(lng).SubItems(4) + "   Id..: " & .Item(lng).Text & "    Notas..: " & txtNotas.Text)
            Call Bitacora("Elimina", "Cargos OP: " & txtOperacion & " Id:" & .Item(lng).Text & "Monto..:" & .Item(lng).SubItems(4))
            
            GLOBALES.gTag2 = 1 'Bandera que indica que se realizó un Ajuste
       
       End If
    Next lng
End With
    
MsgBox "Reversión realizada Satisfactoriamente...", vbInformation
 

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub




