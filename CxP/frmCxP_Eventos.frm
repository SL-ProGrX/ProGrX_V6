VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxP_Eventos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos y Ferias"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10200
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11239
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   9763
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
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "General"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "Label15"
      Item(0).Control(1)=   "Label8"
      Item(0).Control(2)=   "txtCuentaContable"
      Item(0).Control(3)=   "txtCuentaContableDesc"
      Item(0).Control(4)=   "Label12(1)"
      Item(0).Control(5)=   "txtNotas"
      Item(0).Control(6)=   "txtLugar"
      Item(0).Control(7)=   "Label12(0)"
      Item(0).Control(8)=   "txtComision"
      Item(0).Control(9)=   "Label12(2)"
      Item(0).Control(10)=   "dtpInicio"
      Item(0).Control(11)=   "dtpCorte"
      Item(0).Control(12)=   "chkActivo"
      Item(0).Control(13)=   "dtpInicioTime"
      Item(0).Control(14)=   "dtpCorteTime"
      Item(0).Control(15)=   "txtCrdCod"
      Item(0).Control(16)=   "txtCrdDesc"
      Item(0).Control(17)=   "Label12(3)"
      Item(1).Caption =   "Proveedores"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "ShortcutCaption1"
      Item(1).Control(1)=   "lsw"
      Item(1).Control(2)=   "txtFiltro"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activo ?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1035
         Left            =   1680
         TabIndex        =   7
         Top             =   4200
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   1826
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
      Begin XtremeSuiteControls.FlatEdit txtLugar 
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   3600
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
         Height          =   330
         Left            =   1680
         TabIndex        =   11
         ToolTipText     =   "Presione F4"
         Top             =   2640
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaContableDesc 
         Height          =   330
         Left            =   3480
         TabIndex        =   12
         Top             =   2640
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComision 
         Height          =   330
         Left            =   1680
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Height          =   330
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   330
         Left            =   -69880
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicioTime 
         Height          =   330
         Left            =   3120
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   582
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
         Format          =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorteTime 
         Height          =   330
         Left            =   3120
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   582
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
         Format          =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCrdCod 
         Height          =   330
         Left            =   1680
         TabIndex        =   24
         ToolTipText     =   "Presione F4"
         Top             =   3120
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCrdDesc 
         Height          =   330
         Left            =   3480
         TabIndex        =   25
         Top             =   3120
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label12 
         Caption         =   "Linea Crédito"
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
         Left            =   240
         TabIndex        =   26
         Top             =   3120
         Width           =   1335
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   21
         Top             =   420
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Proveedores"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label Label12 
         Caption         =   "Fechas del Evento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Comisión Porc."
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
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Cta. Comisión"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Lugar de Venta"
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
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4200
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Evento"
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
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCxP_Eventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub sbProveedores_List()

If vCodigo = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

vPaso = True

strSQL = "exec spCxP_Eventos_Proveedores_List " & vCodigo & ", '" & txtFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!cod_Proveedor)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = rs!Registro_Fecha & ""
          itmX.SubItems(3) = rs!Registro_Usuario & ""
      
          itmX.Checked = rs!Asignado
      rs.MoveNext
    Loop
    rs.Close

End With

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_Evento from CXP_EVENTOS"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Evento > " & IIf(txtCodigo = "", 0, txtCodigo) & " order by cod_Evento asc"
    Else
       strSQL = strSQL & " where cod_Evento < " & IIf(txtCodigo = "", 0, txtCodigo) & " order by cod_Evento desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Evento)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCxP_Proveedores_Eventos_Asigna " & Item.Text & ", " & vCodigo _
       & ", " & IIf(Item.Checked, 1, 0) & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
   Call Bitacora(rs!Movimiento, rs!Mensaje)
Else
    MsgBox "Este Evento no puede ser modificado porque se encuentra vencido!", vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

If Item.Index = 1 Then
    If vCodigo = 0 Then
       MsgBox "Consulte un Evento Primero...", vbExclamation
       tcMain.Item(0).Selected = True
    Else
       Call sbProveedores_List
    End If
End If

vError:

End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 30

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Nombre", 4200
    .Add , , "Reg.Fecha", 2100, vbCenter
    .Add , , "Reg.Usuario", 2100, vbCenter
End With


 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

vCodigo = 0
txtCodigo.Text = ""

txtCuentaContable.Tag = GLOBALES.gEnlace

txtDescripcion.Text = ""
txtNotas.Text = ""
txtLugar.Text = ""

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


dtpInicioTime.Value = dtpInicio.Value
dtpCorteTime.Value = dtpInicio.Value


txtCuentaContable.Text = ""
txtCuentaContableDesc.Text = ""

txtCrdCod.Text = ""
txtCrdDesc.Text = ""

txtComision.Text = "0"

txtCodigo.Enabled = True

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtDescripcion.SetFocus
      txtCodigo.Enabled = False
      
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select Cod_Evento, Descripcion from cxp_Eventos"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*,isnull(Cta.Descripcion,'') as 'Cuenta_Desc', isnull(Cta.Cod_Cuenta_Mask,'') as 'Cuenta_Mask'" _
       & " , isnull(Crd.Codigo,'') as 'CrdCod', isnull(Crd.Descripcion,'') as 'CrdDesc'" _
       & " from cxp_Eventos E" _
       & " left join CntX_Cuentas Cta on E.Comision_Cuenta = Cta.cod_Cuenta and Cta.cod_contabilidad = " & GLOBALES.gEnlace _
       & " left join Catalogo Crd on E.cod_Linea_Crd = Crd.Codigo" _
       & " where E.cod_Evento = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  
  vEdita = True
  vCodigo = rs!cod_Evento
  txtCodigo.Text = CStr(rs!cod_Evento)
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_finaliza
  
  dtpInicioTime.Value = rs!Fecha_Inicio
  dtpCorteTime.Value = rs!Fecha_finaliza
  
  txtDescripcion.Text = rs!Descripcion & ""
    
  chkActivo.Value = rs!Activo
  
  txtLugar.Text = rs!Lugar_Venta & ""
    
  txtComision.Text = Format(rs!Comision_porc, "Standard")
  
  txtCuentaContable.Text = rs!Cuenta_Mask
  txtCuentaContableDesc.Text = rs!Cuenta_Desc
    
  txtCrdCod.Text = rs!CrdCod
  txtCrdDesc.Text = rs!CrdDesc
    
  txtNotas.Text = rs!Notas & ""
   
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True


'Si Existe Enlace con ContaExpress / Realizar esta verificacion
If txtCuentaContable.Tag = "" Then
  If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtCuentaContable, 0)) Then
     vMensaje = vMensaje & vbCrLf & " - No se especificó una cuenta contable válida..."
  End If
End If

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Evento no es válido ..."
If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - El Lugar del Evento no es válido ..."

If Not IsNumeric(txtComision.Text) Then
    vMensaje = vMensaje & vbCrLf & " - El Porcentaje de la Comisión no es válido!"
Else
    If CCur(txtComision.Text) < 0 Then vMensaje = vMensaje & vbCrLf & " - El Porcentaje de la Comisión no es válido!"
    If CCur(txtComision.Text) > 100 Then vMensaje = vMensaje & vbCrLf & " - El Porcentaje de la Comisión no es válido!"
End If

If txtCrdCod.Text = "" Then
    vMensaje = vMensaje & vbCrLf & " - Indique una Linea de Crédito para el Evento, verifique!"
End If

If dtpInicio.Value > dtpCorte.Value Then
    vMensaje = vMensaje & vbCrLf & " - El Rango de Fechas no es correcto, verifique!"
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim vCuenta As String, vInicio As String, vCorte As String, vEventoId As Long

On Error GoTo vError

vCuenta = fxgCntCuentaFormato(False, txtCuentaContable)
vInicio = Format(dtpInicio.Value, "yyyy-mm-dd") & " " & Format(dtpInicioTime.Value, "HH:MM:SS")
vCorte = Format(dtpCorte.Value, "yyyy-mm-dd") & " " & Format(dtpCorteTime.Value, "HH:MM:SS")

vEventoId = vCodigo
                  

strSQL = "exec spCxP_Eventos_Add " & vEventoId & ", '" & txtDescripcion.Text & "', " & chkActivo.Value _
       & ", '" & vInicio & "', '" & vCorte & "', '" & txtLugar.Text & "', '" & txtNotas.Text _
       & "', " & CCur(txtComision.Text) & ", '" & vCuenta & "', '" & txtCrdCod.Text & "', '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
    vCodigo = rs!Codigo
    txtCodigo.Text = CStr(rs!Codigo)
    Call Bitacora(rs!Movimiento, rs!Mensaje)

    MsgBox "Información guardada satisfactoriamente...", vbInformation
    Call sbConsulta(vCodigo)
Else
   MsgBox rs!Mensaje, vbExclamation
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete cxp_Eventos where cod_Evento = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Evento Id: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Evento"
  gBusquedas.Col2Name = "Descripcion"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_Evento"
  gBusquedas.Orden = "cod_Evento"
  gBusquedas.Consulta = "select cod_Evento,descripcion from cxp_Eventos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub



Private Sub txtCrdCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLugar.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Filtro = " and ACTIVO = 1 and FORMA_PAGO_WEB = 1"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCrdCod.Text = gBusquedas.Resultado
  txtCrdDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCuentaContable_GotFocus()
On Error GoTo vError
If txtCuentaContable.Tag = "S" Then
    txtCuentaContable = fxgCntCuentaFormato(False, txtCuentaContable)
End If
vError:
End Sub

Private Sub txtCuentaContable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCrdCod.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaContable = gCuenta
End If
End Sub

Private Sub txtCuentaContable_LostFocus()
On Error GoTo vError
txtCuentaContable = fxgCntCuentaFormato(True, txtCuentaContable)
txtCuentaContableDesc.Text = fxSIFCCodigos("D", fxgCntCuentaFormato(False, fxgCntCuentaFormato(False, txtCuentaContable)), "cuentas")
vError:
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbProveedores_List
End If
End Sub

Private Sub txtLugar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub txtComision_GotFocus()
On Error GoTo vError
 txtComision.Text = CCur(txtComision.Text)
vError:
End Sub

Private Sub txtComision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaContable.SetFocus
End Sub

Private Sub txtComision_LostFocus()
On Error GoTo vError
 txtComision.Text = Format(CCur(txtComision.Text), "Standard")
vError:
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkActivo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_Evento,descripcion from cxp_Eventos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    dtpInicio.SetFocus
End If
End Sub

