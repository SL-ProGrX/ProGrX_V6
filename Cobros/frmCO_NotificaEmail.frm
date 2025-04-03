VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCO_NotificaEmail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificación vía Email"
   ClientHeight    =   9075
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   12915
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   12735
      _Version        =   1441793
      _ExtentX        =   22463
      _ExtentY        =   9551
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   240
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   10815
      _Version        =   1441793
      _ExtentX        =   19076
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   2080
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmCO_NotificaEmail.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnMail 
      Height          =   615
      Left            =   11160
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Email"
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
      Picture         =   "frmCO_NotificaEmail.frx":0700
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.ComboBox cboCobro 
      Height          =   312
      Left            =   6720
      TabIndex        =   5
      Top             =   1320
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   4560
      TabIndex        =   9
      Top             =   1320
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   12735
      _Version        =   1441793
      _ExtentX        =   22463
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.FlatEdit txtCtaInicia 
      Height          =   315
      Left            =   2880
      TabIndex        =   13
      Top             =   1680
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
      Text            =   "1"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCtaCorte 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   1680
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
      Text            =   "12"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboNotifica 
      Height          =   315
      Left            =   9120
      TabIndex        =   16
      Top             =   1320
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   12240
      TabIndex        =   18
      Top             =   2040
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1085
      _ExtentY        =   661
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
      Picture         =   "frmCO_NotificaEmail.frx":0F1D
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   9120
      TabIndex        =   17
      Top             =   1080
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Notificación para:"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   12735
      _Version        =   1441793
      _ExtentX        =   22463
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Casos encontrados para notificar Atraso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuotas atrasadas entre"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   3
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado Persona"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   612
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11451
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Notificación de Aviso de Cuotas por Vencer"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   6720
      TabIndex        =   7
      Top             =   1080
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Cobro"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Institución/Empresa"
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
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13092
   End
End
Attribute VB_Name = "frmCO_NotificaEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem



Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnMail_Click()
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass
strSQL = ""

With lsw.ListItems

    For i = 1 To .Count
      If .Item(i).Checked Then
        strSQL = strSQL & Space(10) & "exec spSys_Notifica_Cobros_CtaXVencer '" & RTrim(.Item(i).Text) & "','R', '" & glogon.Usuario & "'"
      End If
    
        'Procesa Lote
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    
    Next i

'Procesa Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

End With

Me.MousePointer = vbDefault
MsgBox "Notificaciones Enviadas Satisfactoriamente!", vbInformation


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboNotifica_Click()


With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3800
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Empresa", 3800, vbCenter
    .Add , , "Email", 4800
    
Select Case Mid(cboNotifica.Text, 1, 1)
    Case "D", "F"
        .Add , , "Mora Total", 2100, vbRightJustify
        .Add , , "Mora Cuotas", 2100, vbCenter
    
    Case "A"

        .Add , , "Cta Obrero Pend.", 2100, vbCenter
        .Add , , "Cta Patronal Pend.", 2100, vbCenter
 End Select
End With


End Sub

Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub Form_Activate()
vModulo = 4

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboCobro.Clear
cboCobro.AddItem "Todos"
cboCobro.AddItem "Planilla"
cboCobro.AddItem "Cajas"
cboCobro.Text = "Todos"

cboNotifica.Clear
cboNotifica.AddItem "Deudores"
cboNotifica.AddItem "Fiadores"
cboNotifica.AddItem "Ahorros"
cboNotifica.Text = "Deudores"


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub




Private Sub sbBuscar()
Dim pInstitucion As Long, pEstado As String, pTipoCobro As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

pInstitucion = 0
If cboInstitucion.Text <> "TODOS" Then
   pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

pEstado = "T"
If cboEstado.Text <> "TODOS" Then
   pEstado = cboEstado.ItemData(cboEstado.ListIndex)
End If

pTipoCobro = Mid(cboCobro.Text, 1, 1)

strSQL = "exec spCbr_Consulta_CtaXVencer " & pInstitucion & ",'" & pEstado & "','" & pTipoCobro & "'"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = rs!EstadoDesc
     itmX.SubItems(3) = rs!InstitucionDesc
     itmX.SubItems(4) = rs!Email
     
     itmX.Checked = chkTodos.Value
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - (lsw.Left + 300)
lsw.Height = Me.Height - (lsw.top + 450)


End Sub

Private Sub sbInicializa()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_ESTADO as 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From AFI_ESTADOS_PERSONA"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)

strSQL = "select COD_INSTITUCION as 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From INSTITUCIONES ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub
