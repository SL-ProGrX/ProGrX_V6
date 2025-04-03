VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_DocumentosDup 
   Caption         =   "Consulta de Documentos Duplicados"
   ClientHeight    =   7404
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8880
   Icon            =   "frmTES_DocumentosDup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7404
   ScaleWidth      =   8880
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5412
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   8652
      _Version        =   1245187
      _ExtentX        =   15261
      _ExtentY        =   9546
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
      Height          =   672
      Left            =   7440
      TabIndex        =   5
      Top             =   1080
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_DocumentosDup.frx":030A
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDoc 
      Height          =   312
      Left            =   5160
      TabIndex        =   8
      Top             =   360
      Width           =   3612
      _Version        =   1245187
      _ExtentX        =   6371
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   5160
      TabIndex        =   9
      Top             =   1080
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   312
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   372
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Todos los duplicados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Emitido (Corte)"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Emitido (Inicio)"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
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
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_DocumentosDup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnBuscar_Click()
  Call sbConsulta
End Sub

Private Sub cbo_Click()
If vPaso Then Exit Sub

Call sbTesTiposDocsCargaCbo(cboDoc, cbo.ItemData(cbo.ListIndex))

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9

vPaso = True
    Call sbTesBancoCargaCboAccesoGeneral(cbo)
vPaso = False

With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Solicitud", 1800
    .Add , , "[Id] Cuenta", 1200, vbCenter
    .Add , , "Documento", 2100
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Beneficiario", 4500
    .Add , , "Asiento?", 1200, vbCenter
End With


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call cbo_Click

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Me.Width = 9048

End Sub

Private Sub Form_Resize()

On Error Resume Next

lsw.Width = Me.Width - 450
lsw.Height = Me.Height - 2600


imgBanner.Width = Me.Width

End Sub



Private Sub lsw_DblClick()
If lsw.ListItems.Count <= 0 Then Exit Sub

Dim frm As Form

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(lsw.SelectedItem.Text)
     Exit For
   End If
 Next frm
End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

     strSQL = "select nsolicitud,id_banco,ndocumento,monto,fecha_emision,beneficiario,estado_asiento" _
            & " from Tes_Transacciones where fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and id_banco = " & cbo.ItemData(cbo.ListIndex) _
            & " and tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"
     If Len(Trim(txtDocumento)) > 0 Then
       strSQL = strSQL & " and ndocumento = '" & Trim(txtDocumento) & "'"
     Else
       strSQL = strSQL & " and ndocumento in(select ndocumento" _
            & " from Tes_Transacciones where fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and id_banco = " & cbo.ItemData(cbo.ListIndex) _
            & " and tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'" _
            & " group by ndocumento having count(*) > 1)"
     End If
     
     lsw.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
          itmX.SubItems(1) = rs!id_Banco
          itmX.SubItems(2) = Trim(rs!nDocumento)
          itmX.SubItems(3) = Format(rs!Monto, "Standard")
          itmX.SubItems(4) = Format(rs!Fecha_Emision, "yyyy/mm/dd")
          itmX.SubItems(5) = rs!Beneficiario
          itmX.SubItems(6) = rs!estado_asiento
      rs.MoveNext
     Loop
     rs.Close


Me.MousePointer = vbDefault

End Sub
