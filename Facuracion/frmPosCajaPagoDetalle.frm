VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmPosCajaPagoDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle del Pago"
   ClientHeight    =   5565
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4692
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4572
      _Version        =   1310723
      _ExtentX        =   8064
      _ExtentY        =   8276
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnAgregar 
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   3000
      Width           =   735
      _Version        =   1310723
      _ExtentX        =   1296
      _ExtentY        =   873
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   16
      Picture         =   "frmPosCajaPagoDetalle.frx":0000
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9720
      Top             =   600
   End
   Begin XtremeSuiteControls.FlatEdit txtReferencia 
      Height          =   312
      Left            =   6960
      TabIndex        =   5
      Top             =   1440
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   312
      Left            =   6960
      TabIndex        =   3
      Top             =   5040
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalRecaudo 
      Height          =   312
      Left            =   6960
      TabIndex        =   2
      Top             =   4680
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbono 
      Height          =   672
      Left            =   6960
      TabIndex        =   4
      Top             =   2280
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   1185
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.FlatEdit txtDisponible 
      Height          =   312
      Left            =   6960
      TabIndex        =   6
      Top             =   1920
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtFormaPago 
      Height          =   312
      Left            =   6960
      TabIndex        =   13
      Top             =   1080
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   5
      Left            =   5160
      TabIndex        =   12
      Top             =   1080
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Forma de Pago:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   4
      Left            =   5160
      TabIndex        =   11
      Top             =   1920
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Disponible:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   3
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Referencia:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   2
      Left            =   5160
      TabIndex        =   9
      Top             =   2280
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Abono:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Top             =   5040
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Diferencia"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   4680
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Total Recaudado"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitle 
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10092
      _Version        =   1310723
      _ExtentX        =   17801
      _ExtentY        =   1080
      _StockProps     =   14
      Caption         =   "Monto: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.44
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
End
Attribute VB_Name = "frmPosCajaPagoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnAgregar_Click()

On Error GoTo vError

If txtDisponible.Tag = "S" Then
   If CCur(txtAbono.Text) > CCur(txtDisponible.Text) Then
      MsgBox "El Monto del Abono es Superior al Disponible del Fondo!", vbExclamation
      Exit Sub
   End If
End If

If CCur(txtAbono) + CCur(txtTotalRecaudo.Text) > gCajas.TicketMonto Then
      MsgBox "El Monto del Abono + Recaudo Total es Superior al Monto a Cobrar!", vbExclamation
      Exit Sub
End If
'
'spPV_Cajas_Mov_Tempo_Registra(@Caja varchar(10), @Ticket varchar(50), @FormaPago int, @Monto dec(12,2)
'            , @Referencia varchar(30), @Usuario varchar(30))

'Registra
strSQL = "exec spPV_Cajas_Mov_Tempo_Registra '" & gCajas.Caja & "','" & gCajas.TicketId & "'," & txtFormaPago.Tag _
       & "," & CCur(txtAbono.Text) & ",'" & txtReferencia.Text & "','" & gCajas.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

'Refresca Datos
Call sbFormasPago_Load

 'Salida Automatica
 If gCajas.TicketAbono = gCajas.TicketMonto Then
    Unload Me
 End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

With lsw.ColumnHeaders
    .Clear
    .Add , , "Forma de Pago", 2500
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "No.Referencia"
    .Add , , "Tipo", 500
End With

scTitle.Caption = "TOTAL: " & Format(gCajas.TicketMonto, "Standard")

txtTotalRecaudo.Text = "0"
txtDiferencia.Text = "0"

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtDisponible.Tag = "N"

txtFormaPago.Tag = Item.Tag
txtFormaPago.Text = Item.Text

txtReferencia.Text = ""
txtDisponible.Text = "N/A"
txtAbono.Text = Item.SubItems(1)

If CCur(txtAbono.Text) = 0 Then
    txtAbono.Text = txtDiferencia.Text
End If

'Fondos
If Item.SubItems(3) = "04" Then
    strSQL = "select dbo.fxFnd_Disponible_FP_POS('" & GLOBALES.gTag3 & "') as 'Disponible'"
    Call OpenRecordSet(rs, strSQL)
    txtDisponible.Tag = "S"
    txtDisponible.Text = Format(rs!Disponible, "Standard")
    rs.Close
End If


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbFormasPago_Load

End Sub


Private Sub sbFormasPago_Load()
Dim itmX As ListViewItem, rsTmp As New ADODB.Recordset

On Error GoTo vError

lsw.ListItems.Clear

txtFormaPago.Tag = ""
txtDisponible.Tag = "N"

txtFormaPago.Text = ""
txtReferencia.Text = ""
txtDisponible.Text = "N/A"
txtAbono.Text = "0.00"
txtTotalRecaudo.Text = "0.00"

strSQL = "exec spPV_Cajas_Mov_Tempo '" & gCajas.Caja & "','" & gCajas.TicketId & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!FormaPagoDesc)
     itmX.Tag = rs!Cod_Forma_Pago
     itmX.SubItems(1) = Format(rs!Monto, "Standard")
     itmX.SubItems(2) = rs!Referencia
     itmX.SubItems(3) = rs!Clasificacion

  If txtFormaPago.Tag = "" Then
        txtFormaPago.Tag = rs!Cod_Forma_Pago
        txtFormaPago.Text = rs!FormaPagoDesc
        
        txtReferencia.Text = rs!Referencia
        
        txtDisponible.Text = "N/A"
        txtAbono.Text = Format(rs!Monto, "Standard")

        'Fondos: Disponible
        If rs!Clasificacion = "04" Then
        
            strSQL = "select dbo.fxFnd_Disponible_FP_POS('" & GLOBALES.gTag3 & "') as 'Disponible'"
            
            Call OpenRecordSet(rsTmp, strSQL)
            
            txtDisponible.Tag = "S"
            txtDisponible.Text = Format(rsTmp!Disponible, "Standard")
            rsTmp.Close
    
        End If
  
  End If

  txtTotalRecaudo.Text = Format(CCur(txtTotalRecaudo.Text) + rs!Monto, "Standard")
  rs.MoveNext
Loop
rs.Close

gCajas.TicketAbono = CCur(txtTotalRecaudo.Text)

txtDiferencia.Text = Format(gCajas.TicketMonto - CCur(txtTotalRecaudo.Text), "Standard")

If CCur(txtAbono.Text) = 0 Then
    txtAbono.Text = txtDiferencia.Text
End If


Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub
