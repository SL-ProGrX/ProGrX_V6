VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmInvBodegas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bodegas"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   10275
   Begin TabDlg.SSTab ssTab 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmInvBodegas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboAlmacen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtObservacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboEstado"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "gbCuentas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "gbPermisos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Permisos x Bodegas"
      TabPicture(1)   =   "frmInvBodegas.frx":0708
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "vGrid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboTransac"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin XtremeSuiteControls.GroupBox gbPermisos 
         Height          =   1455
         Left            =   120
         TabIndex        =   24
         Top             =   3720
         Width           =   9615
         _Version        =   1441792
         _ExtentX        =   16960
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Permisos adicionales:"
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
         BorderStyle     =   1
         Begin VB.CheckBox chkSalidas 
            Appearance      =   0  'Flat
            Caption         =   "Permitir Salidas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox chkEntradas 
            Appearance      =   0  'Flat
            Caption         =   "Permitir Entradas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkPermisos 
            Appearance      =   0  'Flat
            Caption         =   "Activar Permisos Transaccionales x Usuario"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3960
            TabIndex        =   25
            Top             =   360
            Width           =   5295
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCuentas 
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   9495
         _Version        =   1441792
         _ExtentX        =   16748
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Cuentas Contables"
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
         BorderStyle     =   1
         Begin VB.TextBox txtCtaGastosTF 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   17
            ToolTipText     =   "Presione F4 Para Consultar"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtCtaIngresosTF 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   16
            ToolTipText     =   "Presione F4 Para Consultar"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtCtaInventario 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   15
            ToolTipText     =   "Presione F4 Para Consultar"
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblCtaGastosTF 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   23
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Label Label12 
            Caption         =   "Gasto x Dif. TF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Cta Gastos por Diferencia de Toma Física"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblCtaIngresosTF 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   21
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label Label12 
            Caption         =   "Ingreso x Dif. TF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "Cta Ingresos por Diferencia de Toma Física"
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Inventario"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Cuenta de Inventarios para la Bodega"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblCtaInventario 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   18
            Top             =   360
            Width           =   5535
         End
      End
      Begin VB.ComboBox cboTransac 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "frmInvBodegas.frx":0E14
         Left            =   -73680
         List            =   "frmInvBodegas.frx":0E24
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   6
         Top             =   960
         Width           =   8895
         _Version        =   524288
         _ExtentX        =   15690
         _ExtentY        =   7223
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvBodegas.frx":0E55
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   7200
         TabIndex        =   10
         Top             =   480
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   675
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   7575
         _Version        =   1441792
         _ExtentX        =   13361
         _ExtentY        =   1191
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboAlmacen 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   7575
         _Version        =   1441792
         _ExtentX        =   13361
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
      Begin VB.Label Label8 
         Caption         =   "Almacen"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisBodegas"
                  Text            =   "Listado de Bodegas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "InvBodegas"
                  Text            =   "Inventario x Bodega"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9360
      TabIndex        =   7
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   1095
      _Version        =   1441792
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   480
      Width           =   6375
      _Version        =   1441792
      _ExtentX        =   11239
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmInvBodegas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vPaso As Boolean, vScroll As Boolean

Private Sub cboTransac_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

If Not vPaso Then Exit Sub


Me.MousePointer = vbHourglass

Select Case Mid(cboTransac.Text, 1, 2)
  Case "En" 'Entradas
    strSQL = "select U.nombre,U.descripcion,isnull(C.E_Modifica,0),isnull(C.E_Autoriza,0),isnull(C.E_Procesa,0)"
  Case "Sa" 'Salidas
    strSQL = "select U.nombre,U.descripcion,isnull(C.S_Modifica,0),isnull(C.S_Autoriza,0),isnull(C.S_Procesa,0)"
  Case "Tr" 'Traslados
    strSQL = "select U.nombre,U.descripcion,isnull(C.T_Modifica,0),isnull(C.T_Autoriza,0),isnull(C.T_Procesa,0)"
  Case "To" 'Toma Fisica
    strSQL = "select U.nombre,U.descripcion,isnull(C.F_Modifica,0),0,isnull(C.F_Procesa,0)"
End Select

strSQL = strSQL & " from usuarios U left join PV_BODEGAS_PERMISOS C on U.nombre = C.usuario" _
           & " and C.cod_bodega = '" & txtCodigo & "' Where U.estado = 'A'" _
           & " order by U.nombre asc"


vPaso = False
 Call sbCargaGrid(vGrid, 5, strSQL)
 vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = True

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
    strSQL = "select Top 1 cod_bodega from pv_bodegas" _
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_bodega > '" & txtCodigo.Text & "' order by cod_bodega asc"
    Else
       strSQL = strSQL & " where cod_bodega < '" & txtCodigo.Text & "' order by cod_bodega desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_bodega)
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

Private Sub Form_Activate()
 vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 32
  
 vGrid.AppearanceStyle = fxGridStyle

 cboAlmacen.AddItem "Almacen General"
 cboAlmacen.Text = "Almacen General"

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
  
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla
 
 vPaso = False
 cboTransac.Text = "Entradas"
 vPaso = True
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

vCodigo = ""
txtCodigo = ""

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "InActivo"
cboEstado.Text = "Activo"

txtNombre = ""
txtObservacion = ""
txtCtaInventario = ""
txtCtaIngresosTF = ""
txtCtaGastosTF = ""

lblCtaInventario.Caption = ""
lblCtaIngresosTF.Caption = ""
lblCtaGastosTF.Caption = ""


chkEntradas.Value = vbUnchecked
chkSalidas.Value = vbUnchecked
chkPermisos.Value = vbUnchecked

ssTab.Tab = 0
vGrid.MaxRows = 0


End Sub





Private Sub SSTab_Click(PreviousTab As Integer)
If ssTab.Tab = 1 Then Call cboTransac_Click
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

ssTab.Tab = 0

strSQL = "select * from pv_bodegas where cod_bodega = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_bodega
  txtCodigo = rs!cod_bodega
 
  txtNombre = rs!Descripcion & ""
  txtObservacion = rs!observacion & ""
    
  If rs!Estado = "A" Then
    cboEstado.Text = "Activo"
  Else
    cboEstado.Text = "InActivo"
  End If
  
  txtCtaInventario = fxgCntCuentaFormato(True, rs!cod_cuenta)
  txtCtaIngresosTF = fxgCntCuentaFormato(True, rs!cod_cta_ingresosTF)
  txtCtaGastosTF = fxgCntCuentaFormato(True, rs!cod_cta_gastosTF)
  
  lblCtaInventario.Caption = fxSIFCCodigos("D", rs!cod_cuenta, "cuentas")
  lblCtaIngresosTF.Caption = fxSIFCCodigos("D", rs!cod_cta_ingresosTF, "cuentas")
  lblCtaGastosTF.Caption = fxSIFCCodigos("D", rs!cod_cta_gastosTF, "cuentas")
  
  chkEntradas.Value = rs!permite_entradas
  chkSalidas.Value = rs!permite_salidas
    
  chkPermisos.Value = rs!UTILIZA_PERMISOS

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

'Validar Cuentas Aqui

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update pv_bodegas set descripcion = '" & UCase(Trim(txtNombre)) & "'" _
         & ",observacion = '" & txtObservacion & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',permite_entradas = " & chkEntradas.Value & ",permite_salidas = " & chkSalidas.Value _
         & ",cod_cuenta = '" & fxgCntCuentaFormato(False, txtCtaInventario) & "',cod_cta_ingresosTF = '" _
         & fxgCntCuentaFormato(False, txtCtaIngresosTF) & "',cod_cta_gastosTF = '" & fxgCntCuentaFormato(False, txtCtaGastosTF) _
         & "',UTILIZA_PERMISOS = " & chkPermisos.Value & " where cod_bodega = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Bodega : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into pv_bodegas(cod_bodega,descripcion,observacion,estado,fecha_inclusion" _
          & ",permite_entradas,permite_salidas,cod_cuenta,cod_cta_ingresosTF,cod_cta_gastosTF,UTILIZA_PERMISOS)" _
          & " values('" & vCodigo & "','" & txtNombre & "','" & txtObservacion & "','" _
          & Mid(cboEstado.Text, 1, 1) & "',dbo.MyGetdate()," & chkEntradas.Value & "," & chkSalidas.Value _
          & ",'" & fxgCntCuentaFormato(False, txtCtaInventario) _
          & "','" & fxgCntCuentaFormato(False, txtCtaIngresosTF) _
          & "','" & fxgCntCuentaFormato(False, txtCtaGastosTF) & "'," & chkPermisos.Value & ")"
   
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Bodega: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete pv_bodegas where cod_bodega = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Bodega : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "LisBodegas"
     Call sbInvReportes("Bodegas", "BODEGAS", "Listado", "")
  Case "InvBodegas"
     Call sbInvReportes("InvBodegas", "BODEGAS", "Inventario", "")
End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_bodega"
  gBusquedas.Orden = "cod_bodega"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCodigo, "bodegas")
End Sub

Private Sub txtCtaInventario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaIngresosTF.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaInventario = gCuenta
End If
End Sub

Private Sub txtCtaInventario_LostFocus()
txtCtaInventario = fxgCntCuentaFormato(True, txtCtaInventario)
lblCtaInventario.Caption = fxSIFCCodigos("D", fxgCntCuentaFormato(False, txtCtaInventario), "cuentas")
End Sub

Private Sub txtCtaIngresosTF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaGastosTF.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaIngresosTF = gCuenta
End If
End Sub

Private Sub txtCtaIngresosTF_LostFocus()
txtCtaIngresosTF = fxgCntCuentaFormato(True, txtCtaIngresosTF)
lblCtaIngresosTF.Caption = fxSIFCCodigos("D", fxgCntCuentaFormato(False, txtCtaIngresosTF), "cuentas")
End Sub


Private Sub txtCtaGastosTF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkEntradas.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaGastosTF = gCuenta
End If
End Sub

Private Sub txtCtaGastosTF_LostFocus()
txtCtaGastosTF = fxgCntCuentaFormato(True, txtCtaGastosTF)
lblCtaGastosTF.Caption = fxSIFCCodigos("D", fxgCntCuentaFormato(False, txtCtaGastosTF), "cuentas")
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaInventario.SetFocus
End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not vPaso Then Exit Sub

If col > 2 Then
  vGrid.Row = Row
  vGrid.col = 1
  strSQL = "select isnull(count(*),0) as Existe from PV_BODEGAS_PERMISOS" _
         & " where usuario = '" & vGrid.Text & "' and cod_bodega = '" & txtCodigo & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
    strSQL = "insert PV_BODEGAS_PERMISOS(usuario,cod_bodega,e_modifica,e_autoriza,e_procesa" _
           & ",s_modifica,s_autoriza,s_procesa,t_modifica,t_autoriza,t_procesa,f_modifica,f_procesa)" _
           & " values('" & vGrid.Text & "','" & txtCodigo & "',0,0,0,0,0,0,0,0,0,0,0)"
    Call ConectionExecute(strSQL)
  End If
  rs.Close
    
  vGrid.col = col
  vGrid.Row = Row
  
  Select Case col
    Case 3 'Modifica
        Select Case Mid(cboTransac.Text, 1, 2)
           Case "En"
            strSQL = "update PV_BODEGAS_PERMISOS set e_modifica = " & vGrid.Value
           Case "Sa"
            strSQL = "update PV_BODEGAS_PERMISOS set s_modifica = " & vGrid.Value
           Case "Tr"
            strSQL = "update PV_BODEGAS_PERMISOS set t_modifica = " & vGrid.Value
           Case "To"
            strSQL = "update PV_BODEGAS_PERMISOS set f_modifica = " & vGrid.Value
        End Select
    
    Case 4 'Autoriza
        Select Case Mid(cboTransac.Text, 1, 2)
           Case "En"
            strSQL = "update PV_BODEGAS_PERMISOS set e_autoriza = " & vGrid.Value
           Case "Sa"
            strSQL = "update PV_BODEGAS_PERMISOS set s_autoriza = " & vGrid.Value
           Case "Tr"
            strSQL = "update PV_BODEGAS_PERMISOS set t_autoriza = " & vGrid.Value
           Case "To"
            strSQL = "update PV_BODEGAS_PERMISOS set f_modifica = f_modifica"
        End Select
    
    Case 5 'Procesa
        Select Case Mid(cboTransac.Text, 1, 2)
           Case "En"
            strSQL = "update PV_BODEGAS_PERMISOS set e_procesa = " & vGrid.Value
           Case "Sa"
            strSQL = "update PV_BODEGAS_PERMISOS set s_procesa = " & vGrid.Value
           Case "Tr"
            strSQL = "update PV_BODEGAS_PERMISOS set t_procesa = " & vGrid.Value
           Case "To"
            strSQL = "update PV_BODEGAS_PERMISOS set f_procesa = " & vGrid.Value
        End Select
    
  End Select
  vGrid.col = 1
  strSQL = strSQL & " where cod_bodega = '" & txtCodigo & "' and usuario = '" & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
End If


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


