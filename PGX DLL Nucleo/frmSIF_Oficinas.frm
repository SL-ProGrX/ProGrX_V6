VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSIF_Oficinas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oficinas y Agencias"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   11295
      _Version        =   1441793
      _ExtentX        =   19923
      _ExtentY        =   10186
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Oficinas"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "lblOficina"
      Item(0).Control(2)=   "txtTelefono1"
      Item(0).Control(3)=   "txtTelefono2"
      Item(0).Control(4)=   "Label4(2)"
      Item(0).Control(5)=   "Label4(1)"
      Item(0).Control(6)=   "txtDireccion"
      Item(0).Control(7)=   "Label4(0)"
      Item(0).Control(8)=   "btnEdit"
      Item(1).Caption =   "Personal"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "lswMiembros"
      Item(1).Control(1)=   "Label4(3)"
      Item(1).Control(2)=   "Label4(4)"
      Item(1).Control(3)=   "cboOficina"
      Item(1).Control(4)=   "chkApoyo"
      Item(1).Control(5)=   "chkUsuariosEstado"
      Item(1).Control(6)=   "Label4(6)"
      Item(1).Control(7)=   "txtFiltro"
      Item(2).Caption =   "Historial"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "Label4(5)"
      Item(2).Control(1)=   "txtUsuario"
      Item(2).Control(2)=   "ShortcutCaption1"
      Item(2).Control(3)=   "lswHistorial"
      Begin XtremeSuiteControls.ListView lswHistorial 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   7435
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
      End
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   4215
         Left            =   -68440
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1441793
         _ExtentX        =   16536
         _ExtentY        =   7435
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnEdit 
         Height          =   315
         Left            =   10080
         TabIndex        =   20
         Top             =   3870
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Editar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmSIF_Oficinas.frx":0000
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   11175
         _Version        =   524288
         _ExtentX        =   19711
         _ExtentY        =   5741
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   501
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_Oficinas.frx":05FB
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   975
         Left            =   2400
         TabIndex        =   9
         Top             =   4560
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   1720
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -68440
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9975
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
      Begin XtremeSuiteControls.CheckBox chkApoyo 
         Height          =   255
         Left            =   -62320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Rol de Ejecutivo de Apoyo?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   4
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkUsuariosEstado 
         Height          =   255
         Left            =   -62320
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Muestra Solo los No Asignados?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   4
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   312
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   315
         Left            =   -68920
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Left            =   -68440
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   6
         Left            =   -69760
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Usuario"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Oficinas / Agencias"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   5
         Left            =   -69760
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Usuario: "
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   4
         Left            =   -69760
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Miembros"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   3
         Left            =   -69760
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Oficina"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   0
         Left            =   2400
         TabIndex        =   10
         Top             =   4320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Dirección:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono (1)"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   4920
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono (2)"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption lblOficina 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   3840
         Width           =   11295
         _Version        =   1441793
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "..."
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
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficinas y Agencias"
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
      Width           =   3252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_Oficinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbLlenaOficinas()
Dim i As Integer


vPaso = True

With vGrid
    .MaxCols = 8
    .MaxRows = 1
    
    strSQL = "select * from Sif_Oficinas"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     .Row = .MaxRows
     For i = 1 To 7
        .Col = i
        Select Case i
           Case 1
             .Text = rs!cod_oficina
             .CellTag = rs!Telefono_01
             
             .TextTip = TextTipFixed
             .TextTipDelay = 1000
             .CellNote = "Registro : " & rs!Registro_Usuario & "[" & rs!registro_Fecha & "]"
           
           Case 2
             .Text = rs!Descripcion
             .CellTag = rs!Telefono_02
           Case 3
             .Text = rs!COD_UNIDAD
             .CellTag = rs!DIRECCION
           Case 4
             .Text = rs!Cod_Centro_Costo
           Case 5
             Select Case rs!Tipo
               Case "A"
                 .Text = "Apoyo"
               Case "E"
                 .Text = "Ejecutiva"
               Case "S"
                 .Text = "Segregación"
             End Select
           Case 6
             .Value = rs!Oficina_Omision
           Case 7
             .Value = rs!Estado
        End Select
     Next i
     rs.MoveNext
     .MaxRows = .MaxRows + 1
    Loop
    rs.Close

End With

vPaso = False


End Sub



Private Sub btnEdit_Click()
Call sbGuardaDatosAddOficinas
End Sub

Private Sub cboOficina_Click()

If vPaso Then Exit Sub

If cboOficina.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

vPaso = True


txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

With lswMiembros
 .ListItems.Clear
  
strSQL = "exec spSys_Oficinas_Miembros_Consultas '" & cboOficina.ItemData(cboOficina.ListIndex) _
       & "', '" & txtFiltro.Text & "', " & chkApoyo.Value & ", " & chkUsuariosEstado.Value
  
 Call OpenRecordSet(rs, strSQL)
 
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If rs!Asignado = 1 Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
         
         itmX.SubItems(2) = rs!FECHA_INGRESO & ""
         
      End If
  rs.MoveNext
 Loop
 rs.Close

End With


vPaso = False


Me.MousePointer = vbDefault

End Sub

Private Sub chkApoyo_Click()
 Call cboOficina_Click
End Sub

Private Sub sbGuardaDatosAddOficinas()

If lblOficina.Tag = "" Then Exit Sub

Dim pOficina As String, pOficinaId As String
Dim pTel1 As String, pTel2 As String, pDir As String

pOficina = lblOficina.Caption
pOficinaId = lblOficina.Tag

pTel1 = txtTelefono1.Text
pTel2 = txtTelefono2.Text
pDir = txtDireccion.Text

strSQL = "update SIF_Oficinas set Telefono_01 = '" & txtTelefono1.Text & "', Telefono_02 = '" _
       & txtTelefono2.Text & "',direccion = '" & txtDireccion.Text _
       & "' where cod_oficina = '" & lblOficina.Tag & "'"
Call ConectionExecute(strSQL)

Call sbLlenaOficinas


lblOficina.Caption = pOficina
lblOficina.Tag = pOficinaId

txtTelefono1.Text = pTel1
txtTelefono2.Text = pTel2
txtDireccion.Text = pDir

End Sub

Private Sub chkUsuariosEstado_Click()
 Call cboOficina_Click
End Sub

Private Sub Form_Load()

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle

imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2100
    .Add , , "Nombre", 3200
    .Add , , "Ingreso", 1800, vbCenter
End With

With lswHistorial.ColumnHeaders
    .Clear
    .Add , , "", 100
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Oficina", 1000, vbCenter
    .Add , , "Fec. Ingreso", 1800, vbCenter
    .Add , , "Fec. Salida", 1800, vbCenter
    
    .Add , , "Us. Ingreso", 1800, vbCenter
    .Add , , "Us. Salida", 1800, vbCenter
End With



Call Formularios(Me)
Call RefrescaTags(Me)

tcMain.Item(0).Selected = True

btnEdit.Enabled = vGrid.Enabled

Call sbLlenaOficinas

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from sif_oficinas" _
       & " where cod_oficina = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into sif_oficinas(cod_oficina,descripcion,cod_unidad,cod_centro_costo,Tipo,Oficina_Omision,Estado" _
         & ",Telefono_01,Telefono_02,Direccion,registro_fecha,registro_usuario) values('" _
         & (vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 5
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ",'" & txtTelefono1.Text & "','" & txtTelefono2.Text & "','" & txtDireccion.Text _
         & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Oficina: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update sif_oficinas set descripcion = '" & vGrid.Text & "',cod_unidad = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "',cod_centro_costo = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "',Tipo = '"
 vGrid.Col = 5
 strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',Oficina_Omision = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ",Estado = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & ",Telefono_01 = '" & txtTelefono1.Text & "',Telefono_02 = '" & txtTelefono2.Text _
        & "',Direccion = '" & txtDireccion.Text & "' where cod_oficina = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)
 Call Bitacora("Modifica", "Oficina : " & vGrid.Text)


End If
rs.Close

'Revisa si esta marcada como Oficina por Omision
vGrid.Col = 6
If vGrid.Value = 1 Then
  vGrid.Col = 1
  strSQL = "Update SIF_Oficinas set Oficina_Omision = 0 where cod_oficina not in('" & vGrid.Text & "')"
  Call ConectionExecute(strSQL)
End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or cboOficina.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Oficinas_Miembros_Add '" & cboOficina.ItemData(cboOficina.ListIndex) & "', '" & Item.Text & "', " & chkApoyo.Value _
       & ", '" & glogon.Usuario & "', '" & IIf(Item.Checked = True, "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Oficinas
    Call sbLlenaOficinas
  Case 1 'Personal
    vPaso = True
        strSQL = "select rtrim(cod_oficina) as 'IdX',  rtrim(descripcion) as 'ItmX' from sif_oficinas where estado = 1" _
               & " order by cod_oficina"
        Call sbCbo_Llena_New(cboOficina, strSQL, False, True)
    vPaso = False
    Call cboOficina_Click
  Case 2 'Historial
End Select
End Sub


Private Sub sbHistorial()

Me.MousePointer = vbHourglass

lswHistorial.ListItems.Clear
 
strSQL = "select * from dbo.SIF_OFICINA_MIEMBROS_H" _
       & " where usuario = '" & txtUsuario.Text & "' order by cod_historial desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF

 Set itmX = lswHistorial.ListItems.Add(, , "")
     If rs!Calidad = "T" Then
        itmX.SubItems(1) = "Titular"
     Else
        itmX.ForeColor = vbYellow
        itmX.SubItems(1) = "Apoyo"
     End If
 
     itmX.SubItems(2) = rs!cod_oficina
     itmX.SubItems(3) = rs!FECHA_INGRESO
     itmX.SubItems(4) = rs!Fecha_salida & ""
 
     itmX.SubItems(5) = rs!Usuario_Ingresa
     itmX.SubItems(6) = rs!Usuario_Salida & ""
 
    If Not IsNull(rs!Fecha_salida) Then itmX.ForeColor = vbRed
 
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

End Sub


Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call cboOficina_Click
End Sub

Private Sub txtUsuario_Change()
 lswHistorial.ListItems.Clear
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbHistorial
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 8 Then
    vGrid.Row = Row
    vGrid.Col = 1
    
    lblOficina.Tag = vGrid.Text
    txtTelefono1.Text = vGrid.CellTag
    
    vGrid.Col = 2
    lblOficina.Caption = "Oficina : " & vGrid.Text
    txtTelefono2.Text = vGrid.CellTag
    
    vGrid.Col = 3
    txtDireccion.Text = vGrid.CellTag
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = 7 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
   gBusquedas.Columna = "cod_unidad"
   gBusquedas.Consulta = "select cod_unidad as unidad, descripcion from CntX_Unidades"
   gBusquedas.Orden = "cod_unidad"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = vGrid.ActiveCol
   vGrid.Text = gBusquedas.Resultado
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
   gBusquedas.Columna = "cod_centro_costo"
   gBusquedas.Consulta = "select cod_centro_costo as CentroCosto, descripcion from CNTX_CENTRO_COSTOS"
   gBusquedas.Orden = "cod_centro_costo"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = vGrid.ActiveCol
   vGrid.Text = gBusquedas.Resultado
End If

If KeyCode = vbKeyDelete Then
'    vGrid.MaxRows = vGrid.MaxRows + 1
'    vGrid.InsertRows vGrid.ActiveRow, 1
'    vGrid.Row = vGrid.ActiveRow
End If


End Sub

