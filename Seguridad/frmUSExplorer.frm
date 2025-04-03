VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.TaskPanel.v22.1.0.ocx"
Begin VB.Form frmUS_Explorer 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Explorador"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12705
   HelpContextID   =   1001
   Icon            =   "frmUSExplorer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel tpContabilidad 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12705
      _Version        =   1441793
      _ExtentX        =   22410
      _ExtentY        =   741
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   7
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Exportar"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":15162
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.CheckBox chkContabiliza 
         Height          =   255
         Left            =   11640
         TabIndex        =   14
         Top             =   60
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuarios Contabilizados"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   6
         Left            =   9480
         TabIndex        =   13
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   30
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Estado de Usuarios"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":15A33
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   5
         Left            =   8160
         TabIndex        =   12
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   30
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Permisos"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":16051
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   4
         Left            =   6600
         TabIndex        =   11
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   30
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Otorgados"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":16759
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   3
         Left            =   5280
         TabIndex        =   10
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   30
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Detalle"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Checked         =   -1  'True
         Picture         =   "frmUSExplorer.frx":16E80
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Informes"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":17599
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   30
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Refrescar"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":17CA0
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   0
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Editar"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmUSExplorer.frx":183A0
         ImageAlignment  =   0
      End
   End
   Begin XtremeSuiteControls.ListView lswExplorer 
      Height          =   2175
      Left            =   2760
      TabIndex        =   5
      Top             =   795
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   3836
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
      Appearance      =   17
      ShowBorder      =   0   'False
   End
   Begin MSComctlLib.ImageList imgBarra 
      Left            =   6120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":1899B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":1F1FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":25A5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":2C2C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":32B23
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":39385
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":3FBE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":46449
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":4CCAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":5350D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   6720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":59D6F
            Key             =   "autorizado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":605D1
            Key             =   "restringido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":66E33
            Key             =   "Opcion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":6D695
            Key             =   "link"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":73EF7
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":7A759
            Key             =   "user"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":80FBB
            Key             =   "grupos"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":8781D
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":8E07F
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":948E1
            Key             =   "UserDetalle"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":9B143
            Key             =   "GruposDetalle"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":A19A5
            Key             =   "OpcionesDetalle"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":A8207
            Key             =   "formularios"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":AEA69
            Key             =   "OpcionesNodos"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   5220
      Visible         =   0   'False
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   582
      ButtonWidth     =   2249
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Nodo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refrescar"
            Key             =   "refrescar"
            Object.ToolTipText     =   "refreca la información del arbol"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "reportes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ayuda"
            Key             =   "ayuda"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalle"
            Key             =   "detalle"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otorgados"
            Key             =   "permisos"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Permisos"
            Key             =   "Accesos"
            Object.ToolTipText     =   "Mantenimiento de Accesos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Activos"
            Key             =   "Estado"
            Object.Tag             =   "'A'"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "activos"
                  Text            =   "Usuarios Activos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "inactivos"
                  Text            =   "Usuarios Inactivos"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "todos"
                  Text            =   "Todos los Usuarios"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Elimina"
                  Text            =   "Elimina Permisos de Inactivos"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2160
      Left            =   3360
      ScaleHeight     =   940.557
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   795
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3810
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgExplorer"
      Appearance      =   0
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
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   582
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   582
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   2565
      MousePointer    =   9  'Size W E
      Top             =   825
      Width           =   150
   End
End
Attribute VB_Name = "frmUS_Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vDato As String
Dim mbMoving As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Const sglSplitLimit = 500



Private Function fxRol_Codigo(pRol_Name As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select cod_Rol from US_ROLES where Descripcion = '" & pRol_Name & "'"

Call OpenRecordSet(rsX, strSQL, 1)

If rsX.EOF Then
 fxRol_Codigo = ""
Else
 fxRol_Codigo = rsX!cod_rol
End If

rsX.Close

End Function


Private Function fxUser_Name(pUser_Desc As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select Usuario from US_USUARIOS where Nombre = '" & pUser_Desc & "'"

Call OpenRecordSet(rsX, strSQL, 1)

If rsX.EOF Then
 fxUser_Name = ""
Else
 fxUser_Name = rsX!Usuario
End If

rsX.Close

End Function


Private Function fxIndiceMultiple(xKey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xKey = fxIndiceCodigo(xKey)

blnPaso = True

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xKey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xKey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xKey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xKey, i, 50) '50 dfgdes un default ningun asiento es tan largo

End If

fxIndiceMultiple = strResultado

End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim strOpciones As String

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

Select Case Node.Text
  Case "Roles"
        strSQL = "select Descripcion,cod_Rol from US_ROLES order by cod_empresa,Descripcion"
        Call OpenRecordSet(rs, strSQL, 1)
        
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, rs!Descripcion, "Nodo_Rol_Detalle", False, "0x0" & rs!cod_rol & "R")
         rs.MoveNext
        Loop
        rs.Close
 
 
  Case "Usuarios"
  
        If gPortal.Empresa_Id = 0 Then
            strSQL = "select Usuario,UserID,Nombre from US_Usuarios where estado in(" _
                   & tlbPrincipal.Buttons.Item(11).Tag & ")"
        Else
            strSQL = "select U.Usuario,U.UserID,U.Nombre" _
                   & " from US_USUARIOS U inner join PGX_CLIENTES_USERS C on U.USUARIO = C.USUARIO" _
                   & " and C.COD_EMPRESA = " & gPortal.Empresa_Id _
                   & " where U.ESTADO in(" & Me.tlbPrincipal.Buttons.Item(11).Tag & ")"
        End If
        
        Select Case chkContabiliza.Value
          Case xtpChecked
            strSQL = strSQL & " and Contabiliza = 1"
          Case xtpUnchecked
            strSQL = strSQL & " and Contabiliza = 0"
          Case xtpGrayed
        End Select
        
        strSQL = strSQL & " order by nombre"
     
        Call OpenRecordSet(rs, strSQL, 1)
        
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, UCase(rs!Nombre), "Nodo_User_Detalle", False, "0x0" & rs!UserID & "U")
         rs.MoveNext
        Loop
        rs.Close
  
  Case "Opciones"
        strSQL = "select rtrim(Nombre) as nombre,Modulo from US_modulos order by nombre"
        
        Call OpenRecordSet(rs, strSQL, 1)
        
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, rs!Nombre, "Nodo_Opciones_Detalle", True, "0x0" & rs!Modulo & "M")
         rs.MoveNext
        Loop
        rs.Close
  
  
  Case "Clientes"
        strSQL = "select cod_Empresa,Nombre_Largo from PGX_Clientes order by Nombre_Largo"
        Call OpenRecordSet(rs, strSQL)
        
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, Trim(rs!Nombre_Largo), "Nodo_Cliente_Detalle", False, "0x0" & rs!cod_Empresa & "E")
         rs.MoveNext
        Loop
        rs.Close
  Case Else
     
     Select Case Right(vNode.Key, 1)
        Case "M" 'Despliga Formularios
            strSQL = "select * from US_formularios" _
                      & " where modulo = " & fxIndiceCodigo(vNode.Key) _
                      & " order by descripcion"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
             Call sbCreaNodos(Node.Key, rs!Descripcion, "Nodo_Formularios", True, "0x0" & rs!Modulo & "-" & rs!Formulario & "F")
             rs.MoveNext
            Loop
            rs.Close
            
        Case "F" 'Despliga Opciones
            strSQL = "select cod_Opcion,Opcion,Opcion_descripcion as 'Descripcion'" _
                  & " from US_Opciones where modulo = " & fxIndiceMultiple(vNode.Key, "T") _
                  & " and formulario= '" _
                  & fxIndiceMultiple(Node.Key, "N") & "' order by opcion"
                  
            Call OpenRecordSet(rs, strSQL)
            
            Do While Not rs.EOF
             Call sbCreaNodos(Node.Key, rs!Descripcion, "Nodo_Opciones", False, "0x0" & rs!cod_Opcion & "O")
             rs.MoveNext
            Loop
            rs.Close
     End Select

End Select

End Sub


Private Function fxIndiceCodigo(xKey As String) As String
xKey = Mid(xKey, 4, Len(xKey))
xKey = Mid(xKey, 1, Len(xKey) - 1)
fxIndiceCodigo = xKey
End Function


Private Sub sbMuestraDetalleSubNodos()
Dim vCadena As String

Select Case vNode.Parent
  Case "Roles"
    
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      
      lblTitle(1).Caption = lblTitle(1).Caption + "   - MIEMBROS"
      
      strSQL = "select M.registro_Fecha,U.Usuario, U.Nombre,U.Estado, C.Nombre_Largo" _
             & " from US_ROL_MIEMBROS M inner join US_usuarios U on M.Usuario = U.Usuario" _
             & "  inner join PGX_Clientes C on M.cod_Empresa = C.cod_Empresa" _
             & " where M.cod_Rol = '" & fxIndiceCodigo(vNode.Key) & "'"
             
     If gPortal.Empresa_Id > 0 Then
             strSQL = strSQL & " and M.cod_Empresa = " & gPortal.Empresa_Id
     End If
     
      Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Usuario", 2650
      .ColumnHeaders.Add , , "Nombre", 4450
      .ColumnHeaders.Add , , "Fecha", 2850, vbCenter
      .ColumnHeaders.Add , , "Estado", 1450, vbCenter
      .ColumnHeaders.Add , , "Cliente Link", 3450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Usuario)
           itmX.SubItems(1) = rs!Nombre
           itmX.SubItems(2) = Format(rs!Registro_Fecha, "Long Date")
           itmX.SubItems(3) = IIf(rs!ESTADO = "A", "Activo", "Inactivo")
           itmX.SubItems(4) = rs!Nombre_Largo
       rs.MoveNext
      Loop
       rs.Close
     End With
      
    Else
    
      lblTitle(1).Caption = lblTitle(1).Caption + "   - PERMISOS"
              
              
      strSQL = "select O.*,M.nombre as 'ModuloName',P.estado" _
             & " from US_ROLES R inner join US_ROL_PERMISOS P on R.COD_ROL = P.COD_ROL" _
             & " inner join US_OPCIONES O on P.cod_Opcion = O.cod_Opcion" _
             & " inner join US_Modulos M on O.modulo = M.modulo" _
             & " where R.COD_ROL = '" & fxIndiceCodigo(vNode.Key) & "'" _
             & " Order By M.Modulo,O.formulario"
      
     vCadena = ""
     Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      Do While Not rs.EOF
       If vCadena <> Trim(rs!ModuloName) Then
          vCadena = Trim(rs!ModuloName)
        Set itmX = .ListItems.Add(, , rs!ModuloName)
            itmX.ForeColor = vbBlue
            itmX.Bold = True
       End If
       
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!Formulario)
           itmX.SubItems(1) = rs!Opcion
           itmX.SubItems(2) = rs!Opcion_descripcion
           itmX.SubItems(3) = IIf((rs!ESTADO = "A"), "Autorización", "Restricción")
           If rs!ESTADO = "R" Then itmX.ForeColor = vbRed
       
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    End If
    
    
  Case "Usuarios"
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      
     lblTitle(1).Caption = lblTitle(1).Caption + "   - MIEMBRO DE..."
     strSQL = " select M.*,R.Descripcion,C.Nombre_Largo" _
            & " from US_ROL_MIEMBROS M inner join US_ROLES R on M.cod_Rol = R.cod_Rol" _
            & " inner join PGX_Clientes C on M.cod_Empresa = C.cod_Empresa" _
            & " inner join US_usuarios U on M.Usuario = U.Usuario" _
            & " Where U.userID = " & fxIndiceCodigo(vNode.Key)
            
    If gPortal.Empresa_Id > 0 Then
        strSQL = strSQL & " and C.cod_Empresa = " & gPortal.Empresa_Id
    End If
    
    strSQL = strSQL & " order by C.Nombre_Largo, R.Descripcion"

    Call OpenRecordSet(rs, strSQL, 1)
     With lswExplorer
      .ColumnHeaders.Add , , "Rol", 4450
      .ColumnHeaders.Add , , "Fecha", 3450
      .ColumnHeaders.Add , , "Cliente Link", 3450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Descripcion)
           itmX.SubItems(1) = Format(rs!Registro_Fecha, "Long Date")
           itmX.SubItems(2) = rs!Nombre_Largo
       rs.MoveNext
      Loop
       rs.Close
     End With
      
      
    Else
    
     strSQL = "Select Cli.Nombre_Corto as 'ClienteDesc',O.formulario,O.opcion,O.opcion_descripcion,P.estado,R.descripcion as 'RolName'" _
            & ",M.nombre as 'ModuloName',F.descripcion as 'FormName'" _
            & " from US_ROL_Permisos P inner join US_Opciones O on P.cod_Opcion = O.cod_Opcion" _
            & "  inner join US_ROLES R on R.cod_ROL = P.cod_rol" _
            & "  inner join US_ROL_MIEMBROS Ms on R.cod_ROL = Ms.cod_ROL and Ms.cod_Empresa = " & gPortal.Empresa_Id _
            & " inner join US_Modulos M on O.modulo = M.modulo" _
            & " inner join US_formularios F on O.formulario = F.formulario" _
            & " inner join US_Usuarios U on Ms.Usuario = U.Usuario" _
            & " inner join PGX_Clientes Cli on Ms.Cod_Empresa = Cli.Cod_Empresa" _
            & " where U.UserID = " & fxIndiceCodigo(vNode.Key) _
            & " group by Cli.Nombre_Corto,O.formulario,O.opcion,O.opcion_descripcion,P.estado,R.descripcion,M.nombre,F.descripcion"
     Call OpenRecordSet(rs, strSQL, 1)
     vCadena = ""
     With lswExplorer
      .ColumnHeaders.Add , , "Cliente", 3450
      .ColumnHeaders.Add , , "Modulo", 2450
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      .ColumnHeaders.Add , , "Rol", 4450
      Do While Not rs.EOF
       
       If vCadena <> Trim(rs!ModuloName) Then
          vCadena = Trim(rs!ModuloName)
        Set itmX = .ListItems.Add(, , rs!ModuloName)
            itmX.ForeColor = vbBlue
            itmX.Bold = True
       End If
       
       Set itmX = .ListItems.Add(, , rs!ClienteDesc)
           itmX.SubItems(1) = rs!ModuloName
           itmX.SubItems(2) = rs!FormName
           itmX.SubItems(3) = rs!Opcion
           itmX.SubItems(4) = rs!Opcion_descripcion
           itmX.SubItems(5) = IIf((rs!ESTADO = "A"), "Autorización", "Restricción")
           itmX.SubItems(6) = rs!RolName
           
           If rs!ESTADO = "R" Then itmX.ForeColor = vbRed
       
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    
    End If
  
  Case "Opciones"
 
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
  
      strSQL = "select * from US_opciones where modulo = " & fxIndiceCodigo(vNode.Key) & " order by formulario,opcion"

      Call OpenRecordSet(rs, strSQL, 1)
     With lswExplorer
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Formulario)
           itmX.SubItems(1) = rs!Opcion
           itmX.SubItems(2) = rs!Opcion_descripcion
       rs.MoveNext
      Loop
       rs.Close
     End With

    Else
    
      lblTitle(1).Caption = lblTitle(1).Caption + "   - PERMISOS OTORGADOS"

     strSQL = "O.*,R.Descripcion as 'RolName',P.estado" _
            & " from US_ROL_Permisos P inner join US_Opciones O on P.cod_Opcion = O.cod_Opcion" _
            & " inner join US_ROLES R on R.COD_ROL = P.cod_Rol" _
            & " where O.modulo = " & fxIndiceCodigo(vNode.Key) & " order by O.formulario"
     Call OpenRecordSet(rs, strSQL, 1)
     With lswExplorer
      .ColumnHeaders.Add , , "Otorgado", 2800
      .ColumnHeaders.Add , , "Formulario", 3450
      .ColumnHeaders.Add , , "Opción", 2050
      .ColumnHeaders.Add , , "Descripción", 3450
      .ColumnHeaders.Add , , "Tipo", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Grupo)
           itmX.SubItems(1) = rs!Formulario
           itmX.SubItems(2) = rs!Opcion
           itmX.SubItems(3) = rs!Descripcion
           itmX.SubItems(4) = IIf((rs!ESTADO = "A"), "Autorización", "Restricción")
           If rs!ESTADO = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
       
      strSQL = "select O.*,U.nombre as 'Usuario',P.estado" _
             & " from US_ROL_Permisos P inner join US_Opciones O on P.cod_Opcion = O.cod_Opcion" _
             & "  inner join US_ROL_Miembros M on P.cod_Rol = M.cod_Rol and M.cod_empresa = " & gPortal.Empresa_Id _
             & "  inner join US_Usuarios U on U.usuario = M.usuario" _
             & " where O.modulo = " & fxIndiceCodigo(vNode.Key) & " order by O.formulario"
      Call OpenRecordSet(rs, strSQL, 1)
       
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Usuario)
           itmX.SubItems(1) = rs!Formulario
           itmX.SubItems(2) = rs!Opcion
           itmX.SubItems(3) = rs!Descripcion
           itmX.SubItems(4) = IIf((rs!ESTADO = "A"), "Autorización", "Restricción")
           If rs!ESTADO = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
       
       
     End With
    
    End If

End Select

End Sub

Private Sub sbMuestraDetalle()
Dim strOpciones As String

lswExplorer.ListItems.Clear
lswExplorer.ColumnHeaders.Clear


Select Case vNode.Text
  
  Case "Roles" 'Grupos
    
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      
      strSQL = "Select R.*,isnull(C.Nombre_Largo,'General') as 'ClienteAsociado'" _
             & "  from us_Roles R left join PGX_Clientes C on R.cod_Empresa = C.cod_Empresa" _
            & " order by  R.cod_Empresa, R.descripcion"
      Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Código", 1200
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Activo?", 1450, vbCenter
      .ColumnHeaders.Add , , "Fecha", 1850, vbCenter
      .ColumnHeaders.Add , , "Cliente Link", 1450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!cod_rol)
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = IIf((rs!Activo = 1), "Sí", "No")
           itmX.SubItems(3) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
           itmX.SubItems(4) = rs!ClienteAsociado
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    Else
    
      strSQL = "select R.descripcion as 'RolName',O.*,P.estado,F.descripcion as 'FormName'" _
             & " from US_ROLES R inner join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol" _
             & " inner join US_OPCIONES O on P.cod_Opcion = O.cod_Opcion" _
             & " inner join US_FORMULARIOS F on O.formulario = F.formulario" _
             & " order by  R.cod_Empresa, R.descripcion"
     Call OpenRecordSet(rs, strSQL, 1)
     With lswExplorer
      .ColumnHeaders.Add , , "Rol", 4450
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!RolName)
           itmX.SubItems(1) = rs!FormName
           itmX.SubItems(2) = rs!Opcion
           itmX.SubItems(3) = rs!Opcion_descripcion
           itmX.SubItems(4) = IIf((rs!ESTADO = "A"), "Autorización", "Restricción")
           If rs!ESTADO = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    
    End If
    
  Case "Usuarios"
     If gPortal.Empresa_Id = 0 Then
            strSQL = "select * from US_usuarios where estado in(" & Me.tlbPrincipal.Buttons.Item(11).Tag & ")"
     Else
            strSQL = "select U.*" _
                   & " from US_USUARIOS U inner join PGX_CLIENTES_USERS C on U.USUARIO = C.USUARIO" _
                   & " and C.COD_EMPRESA = " & gPortal.Empresa_Id _
                   & " where U.ESTADO in(" & Me.tlbPrincipal.Buttons.Item(11).Tag & ")"
     End If
     
     Select Case chkContabiliza.Value
       Case xtpChecked
         strSQL = strSQL & " and Contabiliza = 1"
       Case xtpUnchecked
         strSQL = strSQL & " and Contabiliza = 0"
       Case xtpGrayed
     End Select
     
     strSQL = strSQL & " order by nombre"
     Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Estado", 1450
      .ColumnHeaders.Add , , "Nombre", 4450
      .ColumnHeaders.Add , , "Ingreso", 1450
      .ColumnHeaders.Add , , "Ult.Mov", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Usuario)
           itmX.SubItems(1) = IIf((rs!ESTADO = "A"), "Activo", "Inactivo")
           itmX.SubItems(2) = rs!Nombre
           itmX.SubItems(3) = Format(rs!Registro_Fecha, "dd/mm/yyyy")
           itmX.SubItems(4) = Format(rs!Fecha_Mod, "dd/mm/yyyy")
           itmX.Tag = rs!UserID
       rs.MoveNext
      Loop
       rs.Close
     End With
      
  Case "Opciones"
     
     strSQL = "select * from US_modulos order by modulo"
     Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Modulo", 1450
      .ColumnHeaders.Add , , "Descripción", 4450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Modulo)
           itmX.SubItems(1) = IIf(Len(Trim(rs!Descripcion)) = 0, rs!Nombre, rs!Descripcion)
       rs.MoveNext
      Loop
       rs.Close
     End With
 
 
  Case "Clientes"
     
     strSQL = "select * from PGX_Clientes order by Nombre_Largo"
     Call OpenRecordSet(rs, strSQL, 1)
     
     With lswExplorer
      .ColumnHeaders.Add , , "Código", 1450
      .ColumnHeaders.Add , , "Nombre", 5450
      .ColumnHeaders.Add , , "Corto", 2150
      .ColumnHeaders.Add , , "Estado", 1350, vbCenter
      .ColumnHeaders.Add , , "Identificación", 2250
      .ColumnHeaders.Add , , "Susc.Inicio", 1450, vbCenter
      .ColumnHeaders.Add , , "Susc.Vence", 1450, vbCenter
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!cod_Empresa)
           itmX.SubItems(1) = rs!Nombre_Largo & ""
           itmX.SubItems(2) = rs!Nombre_Corto & ""
           itmX.SubItems(3) = IIf((rs!ESTADO = "A"), "Activo", "Inactivo")
           itmX.SubItems(4) = rs!Identificacion
           itmX.SubItems(5) = Format(rs!Suscripcion_Inicial & "", "dd/mm/yyyy")
           itmX.SubItems(6) = Format(rs!Suscripcion_Vence & "", "dd/mm/yyyy")
       rs.MoveNext
      Loop
       rs.Close
     End With
 
  Case "US"
  
     With lswExplorer
      .ColumnHeaders.Add , , "Empresa", 2450
      .ColumnHeaders.Add , , "Servidor", 2450
      .ColumnHeaders.Add , , "B.D.", 2450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
    
       Set itmX = .ListItems.Add(, , GLOBALES.gstrNombreEmpresa)
           itmX.SubItems(1) = glogon.Servidor
           itmX.SubItems(2) = glogon.BaseDatos
           itmX.SubItems(3) = glogon.Usuario
           itmX.SubItems(4) = Format(fxFechaServidor, "dd/mm/yyyy")
     End With
  
  Case Else
  
    Select Case Right(vNode.Key, 1)
      Case "M" 'Muestra Formularios
            strSQL = "select * from US_FORMULARIOS" _
                   & " where modulo = " & fxIndiceCodigo(vNode.Key) _
                   & " order by formulario"
            Call OpenRecordSet(rs, strSQL, 1)
            
            With lswExplorer
             .ColumnHeaders.Add , , "Formulario", 4450
             .ColumnHeaders.Add , , "Descripción", 6450
             Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Formulario)
                  itmX.SubItems(1) = IIf(Len(Trim(rs!Descripcion)) = 0, rs!Formulario, rs!Descripcion)
              rs.MoveNext
             Loop
              rs.Close
            End With
      
      Case "F" 'Muestra Opciones
            strSQL = "select O.cod_Opcion, O.Opcion, O.Opcion_descripcion as 'Descripcion'" _
                  & " from US_OPCIONES O" _
                  & " inner join US_Formularios F on O.Modulo = F.modulo and O.formulario = F.formulario" _
                  & " where O.modulo = " & fxIndiceMultiple(vNode.Key, "T") _
                  & " and F.Formulario = '" & fxIndiceMultiple(vNode.Key, "N") & "'  order by O.opcion"
            Call OpenRecordSet(rs, strSQL, 1)
            
            With lswExplorer
             .ColumnHeaders.Add , , "Opción: ID", 1050
             .ColumnHeaders.Add , , "Opcion: Name", 4450
             .ColumnHeaders.Add , , "Descripción", 6450
             Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!cod_Opcion)
                  itmX.SubItems(1) = rs!Opcion
                  itmX.SubItems(2) = rs!Descripcion
              rs.MoveNext
             Loop
              rs.Close
            End With
      
      
      Case "O" 'Muestra Asignacion
            With lswExplorer
             .ColumnHeaders.Add , , "", 5450
            
            Set itmX = .ListItems.Add(, , ">>> AUTORIZACIONES")
                itmX.Bold = True
                itmX.ForeColor = vbBlue
            Set itmX = .ListItems.Add(, , "")
            
            
            strSQL = "select R.Descripcion" _
                   & " from US_ROLES R inner join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol" _
                   & " where P.cod_Opcion = " & fxIndiceCodigo(vNode.Key) & " and P.estado = 'A'" _
                   & " group by R.Descripcion"
            Call OpenRecordSet(rs, strSQL, 1)
            
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "ROLES")
                    itmX.Bold = True
                    itmX.ForeColor = vbBlue
            End If
             
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , Space(10) & rs!Descripcion)
              rs.MoveNext
            Loop
             rs.Close
            
                  
            strSQL = "Select U.Nombre" _
                   & " from US_ROLES R inner join US_ROL_MIEMBROS M on R.cod_Rol = M.cod_Rol" _
                   & " inner join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol" _
                   & " inner join US_USUARIOS U on M.usuario = U.usuario" _
                   & " where P.cod_Opcion = " & fxIndiceCodigo(vNode.Key) _
                   & "   and P.estado = 'A'" _
                   & " Group by U.Nombre"
                   
            Call OpenRecordSet(rs, strSQL, 1)
            
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "")
                Set itmX = .ListItems.Add(, , "USUARIOS")
                    itmX.Bold = True
                    itmX.ForeColor = vbBlue
            End If
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , Space(10) & rs!Nombre)
              rs.MoveNext
            Loop
            rs.Close
            
            
            Set itmX = .ListItems.Add(, , "")
            Set itmX = .ListItems.Add(, , ">>> RESTRICCIONES")
                itmX.Bold = True
                itmX.ForeColor = vbRed
            Set itmX = .ListItems.Add(, , "")
            
            strSQL = "select R.Descripcion" _
                   & " from US_ROLES R inner join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol" _
                   & " where P.cod_Opcion = " & fxIndiceCodigo(vNode.Key) & " and P.estado = 'R'" _
                   & " group by R.Descripcion"
            Call OpenRecordSet(rs, strSQL, 1)
            
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "ROLES")
                    itmX.Bold = True
                    itmX.ForeColor = vbRed
            End If
             
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Descripcion)
              rs.MoveNext
            Loop
             rs.Close
            
            strSQL = "Select U.Nombre" _
                   & " from US_ROLES R inner join US_ROL_MIEMBROS M on R.cod_Rol = M.cod_Rol" _
                   & " inner join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol" _
                   & " inner join US_USUARIOS U on M.usuario = U.usuario" _
                   & " where P.cod_Opcion = " & fxIndiceCodigo(vNode.Key) _
                   & "   and P.estado = 'R'" _
                   & " Group by U.Nombre"
            Call OpenRecordSet(rs, strSQL, 1)
            
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "")
                Set itmX = .ListItems.Add(, , "USUARIOS")
                    itmX.Bold = True
                    itmX.ForeColor = vbRed
            End If
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Nombre)
              rs.MoveNext
            Loop
            rs.Close
            
            
            End With
            
       Case "E" 'Muestra Contactos del Cliente
'            strSQL = "select O.cod_Opcion, O.Opcion, O.Opcion_descripcion as 'Descripcion'" _
'                  & " from US_OPCIONES O" _
'                  & " inner join US_Formularios F on O.Modulo = F.modulo and O.formulario = F.formulario" _
'                  & " where O.modulo = " & fxIndiceMultiple(vNode.Key, "T") _
'                  & " and F.Formulario = '" & fxIndiceMultiple(vNode.Key, "N") & "'  order by O.opcion"
            
            strSQL = "select * from PGX_Clientes_Contactos" _
                   & " where cod_empresa = " & fxIndiceCodigo(vNode.Key) _
                   & " and Activo = 1"
            Call OpenRecordSet(rs, strSQL, 1)
            
            With lswExplorer
             .ColumnHeaders.Add , , "Nombre", 3050
             .ColumnHeaders.Add , , "Tel.Cel.", 1250
             .ColumnHeaders.Add , , "Tel.Tra.", 1250
             .ColumnHeaders.Add , , "Email (1)", 3250
             .ColumnHeaders.Add , , "Email (2)", 3250
             Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Nombre)
                  itmX.SubItems(1) = rs!Tel_Cell
                  itmX.SubItems(2) = rs!Tel_Trabajo
                  itmX.SubItems(3) = rs!Email_01
                  itmX.SubItems(4) = rs!Email_02
                  
                  itmX.Tag = rs!cod_Contacto
              rs.MoveNext
             Loop
              rs.Close
            End With
      
      Case Else
        Call sbMuestraDetalleSubNodos
     
     End Select
End Select

End Sub

Private Sub ArbolExp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 2 Then
' Call PopupMenu(MDIMenu.mnuAcciones, , x, y)
'End If
End Sub

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next


Set vNode = Node


lblTitle(0).Caption = UCase(vNode.Text)
lblTitle(1).Caption = vNode.FullPath

Call sbMuestraDetalle


Dim itmX As ListViewItem
 Set itmX = lswExplorer.ListItems.Add(, , "_____________")
     itmX.SubItems(1) = "_____________"


 Set itmX = lswExplorer.ListItems.Add(, , "Items: ")
     itmX.SubItems(1) = Format(lswExplorer.ListItems.Count - 2, "###,###0")

End Sub


Private Sub sbReporteGrupos()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Seguridad"
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Reporte='Reporte Al  " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("SegListadoGrupos.rpt")
    .Connect = glogon.ConectRPT
    .PrintReport
End With
 
Me.MousePointer = vbDefault

End Sub


Private Sub sbReporteOpciones()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Seguridad"
 
     .Connect = glogon.ConectRPT
 
     .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(1) = "Reporte='Reporte Al  " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          
          
     .Formulas(3) = "fxUsuario='USER :" & glogon.Usuario & "'"
     .Formulas(4) = "fxServidor='SERVER :" & glogon.Servidor & "'"
     .Formulas(5) = "fxBaseDatos='DATABASE :" & glogon.BaseDatos & "'"
          
          
If tlbPrincipal.Buttons(6).Value = tbrPressed Then
    'Detalle
        .ReportFileName = SIFGlobal.fxPathReportes("SegListadoOpciones.rpt")
Else
    
    'Permisos
        .ReportFileName = SIFGlobal.fxPathReportes("SegListadoOpcionesOtorgadas.rpt")
   
   
   Select Case Right(vNode.Key, 1)
     Case "M" 'Modulo
        .SelectionFormula = "{US_MODULOS.MODULO} = " & fxIndiceCodigo(vNode.Key)
     Case "F"
        .SelectionFormula = "{US_MODULOS.MODULO} = " & fxIndiceMultiple(vNode.Key, "T") _
                          & " and {US_FORMULARIOS.FRMID} = " & fxIndiceMultiple(vNode.Key, "N")
     Case "O"
        .SelectionFormula = "{US_OPCIONES.cod_Opcion} = " & fxIndiceCodigo(vNode.Key)
   End Select
     
End If
     .PrintReport
End With
Me.MousePointer = vbDefault

End Sub



Public Sub sbButtonPopUp(i As Integer)

On Error GoTo vError
'
'Select Case i
' Case 1 'Editar
'   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
' Case 2 'Reportes
'   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(3))
' Case 3 'Permisos
'   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(9))
'End Select

vError:

End Sub

Private Sub sbExportar()
On Error GoTo vError

Me.MousePointer = vbHourglass


Call Excel_Exportar_Lsw(lswExplorer)


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAccion_Click(Index As Integer)
GLOBALES.gstrReporte = "ListadoUsuarios"

Select Case Index
  Case 0 'editar
    Select Case vNode.Text
      Case "Roles"
        frmUS_Roles.Show
      
      Case "Usuarios"
        frmUS_Usuarios.Show
        
      Case "Opciones"
        frmUS_Opciones.Show
      
      Case "Clientes"
        frmPGX_Clientes.Show
        
        
      Case Else
 
       If vNode.Index > 1 Then
           Select Case vNode.Parent
              Case "Roles"
                    gEntidad.Tipo = "R"
                    gEntidad.Rol_Name = vNode.Text
                    gEntidad.Rol_Id = fxIndiceCodigo(vNode.Key)
                
                frmUS_Roles.Show
              Case "Usuarios"
                    gEntidad.Tipo = "U"
                    gEntidad.Usuario = fxUser_Name(vNode.Text)
                    gEntidad.UserID = fxIndiceCodigo(vNode.Key)
                frmUS_Usuarios.Show
              Case "Opciones"
                frmUS_Opciones.Show
            End Select
       End If
    End Select
    
  Case 10 '"Accesos"
       If vNode.Index > 1 Then
           Select Case vNode.Parent
              Case "Roles"
                    gEntidad.Tipo = "R"
                    gEntidad.Rol_Name = vNode.Text
                    gEntidad.Rol_Id = fxIndiceCodigo(vNode.Key)
              
                    frmUS_DerechosNew.Show
                    frmUS_DerechosNew.cmdDeshacer_Click
              Case "Usuarios"
                    gEntidad.Tipo = "U"
                    gEntidad.Usuario = vNode.Text
                    gEntidad.UserID = fxIndiceCodigo(vNode.Key)
                    
                    frmUS_DerechosNew.Show
                    frmUS_DerechosNew.cmdDeshacer_Click
            End Select
       End If
    
  Case 1 'refrescar"
    Call sbRefrescaArbol
    
    
  Case 2 'reportes"
    
    Select Case vNode.Text
      Case "Roles"
        Call sbReporteGrupos
      Case "Usuarios"
        frmUS_ReporteUsuarios.Show
      
      Case "Opciones"
        Call sbReporteOpciones
      Case Else
        If vNode.Index > 1 Then
            Select Case vNode.Parent
              Case "Grupos"
                Call sbReporteGrupos
              Case "Usuarios"
                frmUS_ReporteUsuarios.Show
                frmUS_ReporteUsuarios.txtUsuario = vNode.Text
              ' Case "Opciones"
              Case Else
                 Call sbReporteOpciones
              
            End Select
        End If
    End Select

   
   Case 3, 4 ' "detalle", "permisos"
      lblTitle(0).Caption = vNode.FullPath
      If vNode.Index > 1 Then
         lblTitle(1).Caption = UCase(vNode.Parent) & " : " & UCase(vNode.Text)
      Else
         lblTitle(1).Caption = vNode.Text
      End If
      Call sbMuestraDetalle

    Case 7
        Call sbExportar

End Select

End Sub

Private Sub chkContabiliza_Click()
    Call sbRefrescaArbol
End Sub

Private Sub Form_Load()
vModulo = 13
 
Dim i As Integer

For i = 0 To btnAccion.Count - 1
     btnAccion.Item(i).FlatStyle = True
     btnAccion.Item(i).Top = 30
Next i

Call sbRefrescaArbol


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = True
      Me.WindowState = 1
   End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    
    'set the width
    If x < 3360 Then x = 3360
    If x > (Me.Width - 3360) Then x = Me.Width - 3360
    ArbolExp.Width = x
    imgSplitter.Left = x
    lswExplorer.Left = x + 40
    lswExplorer.Width = Me.Width - (ArbolExp.Width + 240)
    lblTitle(0).Width = ArbolExp.Width
    lblTitle(1).Left = lswExplorer.Left + 20
    lblTitle(1).Width = lswExplorer.Width - 40

    
'    chkContabiliza.Left = (lblTitle(1).Left + lblTitle(1).Width) - (chkContabiliza.Width + 120)
    
    'set the top
    lswExplorer.Top = ArbolExp.Top
    imgSplitter.Top = ArbolExp.Top
    ArbolExp.Height = Me.Height - 1300

    lswExplorer.Height = ArbolExp.Height
    imgSplitter.Height = ArbolExp.Height
End Sub

Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xKey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xKey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xKey
    End If
    
   
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "US", "Root", "Root")
  vNode.Bold = True
  'Crear Arbol Inicial
   If Sys_Portal_Admin_Valid(glogon.Usuario) Then
        Call sbCreaNodos("US", "Roles", "Roles", True)
        Call sbCreaNodos("US", "Usuarios", "user", True)
        Call sbCreaNodos("US", "Clientes", "Clientes", True)
        Call sbCreaNodos("US", "Opciones", "Opcion", True)
   Else
        Call sbCreaNodos("US", "Usuarios", "user", True)
        Call sbCreaNodos("US", "Opciones", "Opcion", True)
   End If
  
  .Nodes(1).Expanded = True
  
     With lswExplorer
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Empresa", 2450
      .ColumnHeaders.Add , , "Servidor", 2450
      .ColumnHeaders.Add , , "B.D.", 2450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
       Set itmX = .ListItems.Add(, , GLOBALES.gstrNombreEmpresa)
           itmX.SubItems(1) = glogon.Servidor
           itmX.SubItems(2) = glogon.BaseDatos
           itmX.SubItems(3) = glogon.Usuario
           itmX.SubItems(4) = Format(fxFechaServidor, "dd/mm/yyyy")
     End With
  
End With

End Sub

Function fxIndice(Str As String) As String
Dim nodX As Node, lng As Long

On Error Resume Next
With ArbolExp
  For lng = 2 To .Nodes.Count
    Set nodX = .Nodes.Item("0x0" & lng)
    If nodX.Text = Str Then
     fxIndice = "0x0" & lng
     Exit Function
    End If
  Next lng
End With
fxIndice = "0"
End Function



Private Sub lswExplorer_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswExplorer.SortKey = ColumnHeader.Index - 1
  If lswExplorer.SortOrder = 0 Then lswExplorer.SortOrder = 1 Else lswExplorer.SortOrder = 0
  lswExplorer.Sorted = True
End Sub

Private Sub lswExplorer_DblClick()

If lswExplorer.ListItems.Count <= 0 Then Exit Sub

Select Case lblTitle.Item(1).Caption
   Case "Root\Roles"
         gEntidad.Tipo = "R"
         gEntidad.Rol_Name = lswExplorer.SelectedItem.SubItems(1)
         gEntidad.Rol_Id = lswExplorer.SelectedItem
     
     frmUS_Roles.Show
   
   Case "Root\Usuarios"
         gEntidad.Tipo = "U"
         If lswExplorer.SelectedItem Is Nothing Then
            gEntidad.Usuario = ""
            gEntidad.UserID = 0
         Else
            gEntidad.Usuario = lswExplorer.SelectedItem
            gEntidad.UserID = lswExplorer.SelectedItem.Tag
         End If
     frmUS_Usuarios.Show
     
   Case "Root\Clientes"
         GLOBALES.gTag = lswExplorer.SelectedItem
     frmPGX_Clientes.Show
     
 End Select

End Sub


Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String

Select Case ButtonMenu.Key
 Case "activos"
    tlbPrincipal.Buttons.Item(11).Caption = "Activos"
    tlbPrincipal.Buttons.Item(11).Tag = "'A'"
    tlbPrincipal.Buttons.Item(11).Image = 8
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra solo los usuarios activos"
    
    Call sbRefrescaArbol
    
 Case "inactivos"
    tlbPrincipal.Buttons.Item(11).Caption = "Inactivos"
    tlbPrincipal.Buttons.Item(11).Tag = "'I'"
    tlbPrincipal.Buttons.Item(11).Image = 9
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra solo los usuarios Inactivos"
    
    Call sbRefrescaArbol
    
 Case "todos"
    tlbPrincipal.Buttons.Item(11).Caption = "Todos"
    tlbPrincipal.Buttons.Item(11).Tag = "'A','I'"
    tlbPrincipal.Buttons.Item(11).Image = 10
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra TODOS los usuarios"
    
    Call sbRefrescaArbol
    
 Case "Elimina"
   Me.MousePointer = vbHourglass
   
'   strSQL = "delete permisos" _
'          & " where tipo = 'U' and nombre in(select UserId from usuarios where estado = 'I')"
'   Call ConectionExecute(strSQL)
'
'   strSQL = "delete ROL_MIEMBROS" _
'          & " where nombre in(select Nombre from usuarios where estado = 'I')"
'   Call ConectionExecute(strSQL)
   
    MsgBox "Actualización realizada satisfactoriamente...", vbInformation
   
   Me.MousePointer = vbDefault
 

End Select


End Sub
