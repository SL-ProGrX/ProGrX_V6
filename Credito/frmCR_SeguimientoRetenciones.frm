VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_SeguimientoRetenciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Refundiciones de Retenciones"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswPrestamos 
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18018
      _ExtentY        =   4895
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
   Begin XtremeSuiteControls.ListView lswRefunde 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   4043
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
   Begin VB.Frame fraRefunde 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   10212
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1455
         Left            =   3120
         TabIndex        =   7
         Top             =   1200
         Width           =   3975
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Cuotas vencidas y atrasadas:"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   360
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCargos 
            Height          =   312
            Left            =   1800
            TabIndex        =   9
            Top             =   1080
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIVA 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   720
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Cargos + Pólizas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   252
            Index           =   9
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label2 
            Caption         =   "Amortización"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.PushButton btnRefunde 
         Height          =   495
         Left            =   7920
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Refunde"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoRetenciones.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   495
         Left            =   9360
         TabIndex        =   13
         Top             =   2640
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoRetenciones.frx":0727
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   312
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   1440
         TabIndex        =   16
         Top             =   1080
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
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
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   315
         Left            =   4920
         TabIndex        =   24
         Top             =   2760
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Total Pendiente"
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
         Index           =   5
         Left            =   3240
         TabIndex        =   25
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
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
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Línea"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
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
         Left            =   3360
         TabIndex        =   18
         Top             =   720
         Width           =   612
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   10212
         _Version        =   1441793
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Datos de la Refundición o Abono a la operación:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8040
      Top             =   360
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Refundiciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label lblDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   8160
      TabIndex        =   2
      Top             =   1404
      Width           =   2052
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   1
      Top             =   1404
      Width           =   1092
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Refundición de Retenciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   5412
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10212
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Operaciones activas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmCR_SeguimientoRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OpARefundir
  Operacion As Long
  Saldo     As Currency
  Amortiza  As Currency
  Cargos    As Currency
  IVA       As Currency
End Type

Dim mRetencion As OpARefundir
Dim curPrimerCuota As Currency, curPoliza As Currency, curInteres As Currency

Private Sub btnCerrar_Click()
Call LimpiaDatos(False)
End Sub

Private Sub btnRefunde_Click()
Call sbRefunde
End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture


curPrimerCuota = 0
curPoliza = 0
curInteres = 0

With lswRefunde.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Mora", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "IVA", 1800, vbRightJustify
End With


With lswPrestamos.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Mora", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "IVA", 1800, vbRightJustify
End With

fraRefunde.top = lswPrestamos.top
fraRefunde.Left = lswPrestamos.Left
fraRefunde.Height = lswPrestamos.Height
fraRefunde.Width = lswPrestamos.Width


End Sub

Private Sub sbCargaRefundiciones()
    Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

    strSQL = "select R.*,isnull(R.Cargos,0) as CargosDef,C.descripcion, 0 as 'IVA' " _
        & " from refunde_retencion R inner join catalogo C on R.codigo = C.codigo" _
        & " Where id_solicitudR = " & Operacion.Operacion

    Call OpenRecordSet(rs, strSQL)

    With lswRefunde
        .ListItems.Clear
        Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
            itmX.SubItems(1) = rs!Codigo
            itmX.SubItems(2) = rs!DESCRIPCION
            itmX.SubItems(3) = Format(rs!Monto, "Standard")
            itmX.SubItems(4) = Format(rs!Mora, "Standard")
            itmX.SubItems(5) = Format(rs!CargosDef, "Standard")
            itmX.SubItems(6) = Format(rs!IVA, "Standard")
            rs.MoveNext
        Loop
    End With
    rs.Close

End Sub

Private Sub sbCargaRetenciones()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

' 1. Saca las retenciones a Plazo, con o sin morosidad
' 2. Saca las retenciones indefinidas pero solo las que tienen morosidad

strSQL = "select R.id_solicitud,R.codigo,C.descripcion,R.amortiza,R.cuota,R.plazo,isnull(V.amortiza,0) as Mora, isnull(V.Cargos,0) as Cargos, 0 as 'IVA'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'S'" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " Where R.proceso <> 'J' and R.estado = 'A' and R.plazo < 900 and R.cedula = '" & Operacion.Cedula & "'" _
       & " UNION " _
       & " Select R.id_solicitud,R.codigo,C.descripcion,R.amortiza,R.cuota,0 as plazo,isnull(V.amortiza,0) as Mora, isnull(V.Cargos,0) as Cargos, 0 as 'IVA'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'S'" _
       & " inner join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " Where R.proceso <> 'J' and R.estado = 'A' and R.plazo >= 900 and R.cedula = '" & Operacion.Cedula & "'"

'Retenciones de Plazo Indefinido > 900

Call OpenRecordSet(rs, strSQL)

With lswPrestamos
  .ListItems.Clear
  Do While Not rs.EOF
     Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
         itmX.SubItems(1) = rs!Codigo
         itmX.SubItems(2) = rs!DESCRIPCION
         
         If rs!Plazo < 900 And rs!Plazo > 0 Then
             itmX.SubItems(3) = Format((rs!Cuota * rs!Plazo) - (rs!Amortiza + rs!Mora), "Standard") 'Saldo - (Mora)
         Else
             itmX.SubItems(3) = 0  'Saldo 0, para Indefinidas / Solo aplica mora
         End If
         itmX.SubItems(4) = Format(rs!Mora, "Standard")
         itmX.SubItems(5) = Format(rs!Cargos, "Standard")
         itmX.SubItems(6) = Format(rs!IVA, "Standard")
        
   rs.MoveNext
  
  Loop
End With
rs.Close

End Sub

Private Function fxExisteRetencion(vOperacion As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from refunde_retencion" _
       & " where id_solicitud = " & vOperacion & " and id_solicitudr=" _
       & Operacion.Operacion

Call OpenRecordSet(rs, strSQL)
    fxExisteRetencion = IIf((rs!Existe = 0), False, True)
rs.Close

End Function

Private Sub LimpiaDatos(Optional vVisible As Boolean = True)


txtCodigo.Text = ""
txtOperacion.Text = ""

txtSaldo.Text = "0.00"
txtAmortizacion.Text = "0.00"
txtCargos.Text = "0.00"
txtIVA.Text = "0.00"
txtTotal.Text = "0.00"


mRetencion.Amortiza = 0
mRetencion.Saldo = 0
mRetencion.Operacion = 0
mRetencion.Cargos = 0
mRetencion.IVA = 0

If vVisible Then
   fraRefunde.Visible = vVisible
Else
   fraRefunde.Visible = vVisible
End If

End Sub



Private Function fxValidaRefundicion() As Boolean
Dim vMensaje As String

fxValidaRefundicion = True
vMensaje = ""

If mRetencion.Operacion = 0 Then vMensaje = vMensaje & "- No se ha seleccionado ninguna operación"

If IsNumeric(txtSaldo.Text) Then
 If txtSaldo.Text > mRetencion.Saldo Then vMensaje = vMensaje & vbCrLf & "- El saldo es mayor que el Original"
 If txtSaldo.Text < 0 Then vMensaje = vMensaje & vbCrLf & "- El saldo no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El saldo no es válido"
End If

If Len(vMensaje) > 0 Then
 fxValidaRefundicion = False
 MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbRefunde()
Dim strSQL As String, curRefundir As Currency

On Error GoTo vError

If fxValidaRefundicion Then

curRefundir = CCur(txtSaldo.Text) + CCur(txtAmortizacion.Text) + CCur(txtCargos.Text)

If curRefundir > CCur(lblDisponible.Caption) Then
  MsgBox "El monto a refundir de la operación es mayor al disponible...", vbCritical
  Exit Sub
End If

If fxExisteRetencion(txtOperacion.Text) Then
  MsgBox "Esta Refundición Se encuentra Registrada VERIFIQUE...", vbInformation
  Exit Sub
Else
  strSQL = "insert refunde_retencion(id_solicitud,codigo,monto,mora,fecha,codigor,id_solicitudr,saldo_anterior,cargos) " _
         & "values(" & txtOperacion.Text & ",'" & txtCodigo.Text & "'," & CCur(txtSaldo.Text) & "," & CCur(txtAmortizacion.Text) _
         & ",dbo.MyGetdate(),'" & Operacion.Codigo & "'," _
         & Operacion.Operacion & "," & mRetencion.Saldo & "," & CCur(txtCargos.Text) & ")"
  Call ConectionExecute(strSQL)
  
  lblDisponible.Caption = CCur(lblDisponible.Caption) - (CCur(txtSaldo.Text) + CCur(txtAmortizacion.Text) + CCur(txtCargos.Text))
  lblDisponible.Caption = Format(lblDisponible, "Standard")
  
  Call sbCargaRefundiciones
  Call LimpiaDatos(False)
  
End If

End If 'Verificacion de OPERACION

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select R.Primer_Cuota,R.Garantia,R.montoapr,R.cuota,R.int,C.convenio,R.cod_destino" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where R.id_solicitud =" & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)

If fxCobraTasaFormaliza(rs!cod_destino & "") Then
  'curInteres = fxInteresesHastaFormalizar(Operacion.FechaDesembolso)
  curInteres = fxInteresesHastaFormalizar(Operacion.FechaDesembolso, , Operacion.PriDeduc, Operacion.DiaPago)
End If

If rs!PRIMER_CUOTA = "S" Then
  curPrimerCuota = rs!Cuota
  If curInteres > 0 Then
     curInteres = fxInteresesDiasPrimerCuota(Operacion.FechaDesembolso, rs!montoapr, rs!Int)
  End If
End If

If rs!Garantia <> "F" And rs!Convenio = "N" Then curPoliza = fxCuotaPolizaVida(rs!montoapr)
rs.Close
    
Me.Caption = "Refundiciones Operación : " & Operacion.Operacion

lblDisponible.Caption = Format(Operacion.MontoAprobado _
                      - (fxMontoEnGeneral(Operacion.Operacion) _
                         + curInteres + curPrimerCuota + curPoliza) _
                      , "Standard")
                      
Call sbCargaRefundiciones
Call sbCargaRetenciones

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswPrestamos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

txtOperacion.Text = Item.Text
txtCodigo.Text = Item.SubItems(1)
txtSaldo.Text = CCur(Item.SubItems(3)) 'El saldo va restado con la mora
txtAmortizacion.Text = CCur(Item.SubItems(4)) 'Solo la mora
txtCargos.Text = CCur(Item.SubItems(5))
 
txtIVA.Text = CCur(Item.SubItems(6))
 
mRetencion.Operacion = txtOperacion.Text
mRetencion.Amortiza = CCur(txtAmortizacion.Text)
mRetencion.Saldo = CCur(txtSaldo.Text)
mRetencion.Cargos = CCur(txtCargos.Text)
mRetencion.IVA = CCur(txtIVA.Text)


txtTotal.Text = mRetencion.Cargos + mRetencion.Amortiza + mRetencion.Saldo + mRetencion.IVA


fraRefunde.Visible = True
  
vError:

End Sub

Private Sub lswRefunde_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError


strSQL = "delete refunde_retencion where id_solicitud = " & Item.Text _
       & " and id_solicitudr = " & Operacion.Operacion
Call ConectionExecute(strSQL)

lblDisponible.Caption = CCur(lblDisponible.Caption) + (CCur(Item.SubItems(3)) + CCur(Item.SubItems(4)) + CCur(Item.SubItems(5)) + CCur(Item.SubItems(6)))
lblDisponible.Caption = Format(lblDisponible, "Standard")

Call sbCargaRefundiciones

vError:
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub


