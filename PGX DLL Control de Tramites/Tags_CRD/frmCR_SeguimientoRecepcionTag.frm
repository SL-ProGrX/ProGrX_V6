VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_SeguimientoRecepcionTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción Documentación Créditos"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   11520
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":D1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":D2FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRecepcionTag.frx":13B5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Recepción - Devolución"
      TabPicture(0)   =   "frmCR_SeguimientoRecepcionTag.frx":1A3C0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblOperacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ImageList1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tlbAplicar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lswOperaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOperacion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAgregar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "PrgBar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optRecepcion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optDevolucion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "frmCR_SeguimientoRecepcionTag.frx":1A3DC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optPendRecepcion"
      Tab(1).Control(1)=   "optPendDevolucion"
      Tab(1).Control(2)=   "tlbBuscar"
      Tab(1).Control(3)=   "vGrid"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Label1(0)"
      Tab(1).Control(6)=   "Image2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Consultas"
      TabPicture(2)   =   "frmCR_SeguimientoRecepcionTag.frx":1A3F8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboUsuario"
      Tab(2).Control(1)=   "txtOperacionBuscar"
      Tab(2).Control(2)=   "vGridConsulta"
      Tab(2).Control(3)=   "dtpFInicio"
      Tab(2).Control(4)=   "dtpFFin"
      Tab(2).Control(5)=   "tlbReportes"
      Tab(2).Control(6)=   "Label1(14)"
      Tab(2).Control(7)=   "Label7"
      Tab(2).Control(8)=   "Line1"
      Tab(2).Control(9)=   "Label6"
      Tab(2).Control(10)=   "Image4"
      Tab(2).Control(11)=   "Label5"
      Tab(2).Control(12)=   "Image3"
      Tab(2).ControlCount=   13
      Begin VB.ComboBox cboUsuario 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -70440
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton optPendRecepcion 
         Caption         =   "Recepción"
         Height          =   495
         Left            =   -71640
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optPendDevolucion 
         Caption         =   "Devolución"
         Height          =   495
         Left            =   -70200
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtOperacionBuscar 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -74160
         TabIndex        =   14
         Top             =   2040
         Width           =   2775
      End
      Begin VB.OptionButton optDevolucion 
         Caption         =   "Devolución"
         Height          =   495
         Left            =   6360
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optRecepcion 
         Caption         =   "Recepción"
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   7320
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtOperacion 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   720
         Width           =   2775
      End
      Begin MSComctlLib.ListView lswOperaciones 
         Height          =   5775
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   10186
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Operación"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Línea"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cédula"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Oficina"
            Object.Width           =   8114
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   120
         TabIndex        =   4
         Top             =   7200
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   1005
         ButtonWidth     =   2461
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Etiqueta"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Opciones"
               Key             =   "Opciones"
               Object.ToolTipText     =   "Opciones"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Limpiar"
                     Text            =   "Limpiar"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Eliminar"
                     Text            =   "Eliminar Crédito"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -66240
         TabIndex        =   6
         Top             =   600
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   7
         Top             =   1200
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   11456
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
         MaxCols         =   486
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoRecepcionTag.frx":1A414
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoRecepcionTag.frx":1AB17
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoRecepcionTag.frx":21379
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoRecepcionTag.frx":27BDB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridConsulta 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   15
         Top             =   2640
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   8493
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
         MaxCols         =   486
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoRecepcionTag.frx":2E43D
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpFInicio 
         Height          =   330
         Left            =   -74160
         TabIndex        =   16
         ToolTipText     =   "Fecha Inicio Búsqueda"
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   293797891
         CurrentDate     =   40361
      End
      Begin MSComCtl2.DTPicker dtpFFin 
         Height          =   330
         Left            =   -72240
         TabIndex        =   17
         ToolTipText     =   "Fecha Fin Búsqueda"
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   293863427
         CurrentDate     =   40361
      End
      Begin MSComctlLib.Toolbar tlbReportes 
         Height          =   330
         Left            =   -67440
         TabIndex        =   19
         Top             =   960
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Inventario"
                     Text            =   "Movimientos"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Recepcion"
                     Text            =   "Recepción"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Devolucion"
                     Text            =   "Devolución"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   14
         Left            =   -70440
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Consulta por Operación"
         Height          =   255
         Left            =   -74160
         TabIndex        =   21
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -74760
         X2              =   -63840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "Reportes"
         Height          =   255
         Left            =   -74160
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmCR_SeguimientoRecepcionTag.frx":2E9DC
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72600
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmCR_SeguimientoRecepcionTag.frx":2EBE6
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Movimiento"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Operaciones Pendientes de:"
         Height          =   375
         Left            =   -74040
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   -74160
         TabIndex        =   9
         Top             =   480
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmCR_SeguimientoRecepcionTag.frx":2EDFF
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmCR_SeguimientoRecepcionTag.frx":2F010
         Top             =   480
         Width           =   480
      End
      Begin VB.Label LblOperacion 
         Caption         =   "Operación"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCR_SeguimientoRecepcionTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem
Dim mTagRecepcion As String, mTagDevolucion As String

Private Sub sbParametrosTags()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    '' Busca el parámetro del tag de recepción
    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '28'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRecepcion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 28 en la base de datos"
    End If
    rs.Close
    
    If Not mTagRecepcion = Empty Then
    
        strSQL = "select COUNT(*) FROM CRD_TAGS where TAG_CODIGO = '" & mTagRecepcion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    '' Busca el parámetro del tag de devolución
    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '29'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 29 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) FROM CRD_TAGS where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

    If Not IsNumeric(txtOperacion) Then
        Exit Sub
    End If
    
    If fxValidaNoDuplicados = True Then
        MsgBox "La operación se ya fue digitada"
        txtOperacion.Text = Empty
        txtOperacion.SetFocus
        Exit Sub
    End If
    
    'Valida no agregar en forma mismo tag en forma consecutiva
    If optRecepcion.Value Then
        strSQL = "SELECT dbo.fxCrdOperacionValidaTagRev(" & Trim(txtOperacion) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "')"
    Else
        strSQL = "SELECT dbo.fxCrdOperacionValidaTagRev(" & Trim(txtOperacion) & ",'" & Trim(mTagDevolucion) & "','" & Trim(mTagRecepcion) & "')"
    End If
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs.Fields(0) = 1 Then
            If optRecepcion.Value Then
                MsgBox "No es posible registrar en forma consecutiva dos recepciones en la operación " & txtOperacion.Text
            Else
                MsgBox "No es posible registrar en forma consecutiva dos devoluciones en la operación " & txtOperacion.Text
            End If
            txtOperacion.Text = Empty
            rs.Close
            Exit Sub
        End If
    End If
    rs.Close
    
    strSQL = "SELECT R.ID_SOLICITUD,R.CODIGO,R.CEDULA,R.FECHAFORF,isnull(O.DESCRIPCION,'') as DESCRIPCION FROM REG_CREDITOS R LEFT JOIN SIF_OFICINAS O ON R.COD_OFICINA_R = O.COD_OFICINA WHERE R.ID_SOLICITUD = " & Trim(txtOperacion)
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswOperaciones.ListItems.Add(, , rs!id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!Cedula
        itmX.SubItems(3) = rs!Descripcion
    End If
    rs.Close

    txtOperacion.Text = Empty
    txtOperacion.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbAplicarRecepcion()
Dim i As Integer, strSQL As String

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar la etiqueta en las operaciones", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If optRecepcion.Value = True Then
    If mTagRecepcion = Empty Then
        MsgBox "No se puede realizar el proceso no está definido la etiqueta de recepción"
        Exit Sub
    End If
Else
    If mTagDevolucion = Empty Then
        MsgBox "No se puede realizar el proceso no está definido la etiqueta de devolución"
        Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

PrgBar.Max = lswOperaciones.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswOperaciones.ListItems

For i = 1 To .Count

    If optRecepcion.Value = True Then
    
''       Se pasa al sp al insertar el tag
'        ''Cambia a ANALISTAS_RECEPCION EN REG_CREDITOS A 'R'
'        strSQL = "update REG_CREDITOS SET ANALISTAS_RECEPCION = 'R' WHERE ID_SOLICITUD = " & Trim(.Item(I).Text)
'        Call ConectionExecute(strSQL)
        
        Call sbCrdOperacionTags(.Item(i).Text, .Item(i).SubItems(1), mTagRecepcion, "", "Recibida la documentación de la operación")
    Else

''       Se pasa al sp al insertar el tag
'        ''Cambia a ANALISTAS_RECEPCION EN REG_CREDITOS A 'D'
'        strSQL = "update REG_CREDITOS SET ANALISTAS_RECEPCION = 'D' WHERE ID_SOLICITUD = " & Trim(.Item(I).Text)
'        Call ConectionExecute(strSQL)
    
        Call sbCrdOperacionTags(.Item(i).Text, .Item(i).SubItems(1), mTagDevolucion, "", "Devolución de la documentación de la operación")
    End If

    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso concluido con éxito...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lswOperaciones.ListItems.Count

        If lswOperaciones.ListItems(i).Text = Trim(txtOperacion.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub cmdAgregar_Click()
    Call sbCargaOperacion
End Sub

Private Sub sbLimpiarDatosCreditos(ByVal Todo As Boolean)

    If Todo = True Then
        txtOperacion.Text = Empty
    End If
    
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
 vModulo = 8
 Call Formularios(Me)
    
    SSTab.Tab = 0
    Call sbParametrosTags
    dtpFInicio.Value = fxFechaServidor
    dtpFFin.Value = dtpFInicio.Value
    vGrid.MaxRows = 0
    

 Call RefrescaTags(Me)
End Sub

Private Sub lswOperaciones_DblClick()
    If lswOperaciones.ListItems.Count > 0 Then
        If lswOperaciones.SelectedItem.Index > 0 Then
            If MsgBox("Desea eliminar el crédito " & lswOperaciones.SelectedItem, vbYesNo) = vbYes Then
                lswOperaciones.ListItems.Remove (lswOperaciones.SelectedItem.Index)
            End If
        End If
    End If
End Sub



Private Sub optDevolucion_Click()
    lswOperaciones.ListItems.Clear
End Sub

Private Sub optRecepcion_Click()
    lswOperaciones.ListItems.Clear
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case SSTab.Tab
    Case 0
        optRecepcion.Value = True
    Case 1
        optPendRecepcion.Value = True
    Case 2
        vGridConsulta.MaxRows = 0
        vGridConsulta.MaxCols = 3
        txtOperacionBuscar.Text = Empty
        txtOperacionBuscar.SetFocus
        
        Call sbCargarUsuarios
    End Select
End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "APLICAR"
        Call sbAplicarRecepcion
    End Select
End Sub

Private Sub tlbAplicar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
    Case "LIMPIAR"
        lswOperaciones.ListItems.Clear
    Case "ELIMINAR"
        If lswOperaciones.ListItems.Count > 0 Then
            If lswOperaciones.SelectedItem.Index <> 0 Then
                lswOperaciones.ListItems.Remove (lswOperaciones.SelectedItem.Index)
            End If
        End If
    End Select
End Sub

Private Sub sbCargarListaSolicitudes(Optional ByVal Num_Operacion As String = Empty)
' Carga Lista de operaciones
    Dim strSQL As String, BancosSeleccionados As String, Estado As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    
    Me.MousePointer = vbHourglass
    vGrid.SetFocus
    dtpFInicio.Refresh
    dtpFFin.Refresh
    
    If optPendRecepcion.Value = True Then
        Estado = "N"
    Else
        Estado = "D"
    End If
        
    strSQL = "select R.ID_SOLICITUD,R.FECHAFORP,R.CEDULA,S.NOMBRE,R.CODIGO,ISNULL(O.DESCRIPCION,''),R.MONTOSOL, R.USERFOR," _
            & " isnull(T.REGISTRO_USUARIO,'') as USUARIO_REVISION,ISNULL(RA.remesa,0) AS REMESA, ISNULL(RE.USUARIO,'') AS USUARIO_REMESA " _
            & " from reg_creditos R " _
            & " inner join SOCIOS S on S.CEDULA = R.CEDULA " _
            & " left join SIF_OFICINAS O on R.COD_OFICINA_R = O.COD_OFICINA " _
            & " LEFT JOIN CATALOGO C ON R.CODIGO = C.CODIGO " _
            & " left join CRD_OPERACION_TAGS T on r.id_solicitud = t.id_solicitud and t.TAG_CODIGO = 'S10' " _
            & " left join CRD_REMESA_ASG RA on R.id_solicitud = RA.id_solicitud " _
            & " left join CRD_REMESAS RE on RE.REMESA = RA.REMESA " _
            & " where R.ESTADOSOL = 'F' and R.REFERENCIA IS NULL and isnull(R.ANALISTAS_RECEPCION,'N') = '" & Estado & "'" _
            & " and R.CODIGO NOT IN (select COD_PLAN from FND_PLANES) AND C.RETENCION = 'N'" _
            & " order by R.ID_SOLICITUD"
                      
    Call sbCargaGrid(vGrid, 11, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub


Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "BUSCAR"
        Call sbCargarListaSolicitudes
    Case "IMPRIMIR"
        Call sbReportes("Pendientes")
End Select
End Sub





Private Sub tlbReportes_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
    Case "INVENTARIO"
        Call sbReportes("Inventario")
    Case "RECEPCION"
        Call sbReportes("Recepcion")
    Case "DEVOLUCION"
        Call sbReportes("Devolucion")
End Select
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargaOperacion
    End If
End Sub

Private Sub sbReportes(pTipo As String)
Dim vFecha As Date

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = False
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"

 .Connect = glogon.ConectRPT
                
    Select Case pTipo
      Case "Inventario"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecepcionDocs.rpt")
      Case "Pendientes"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecepcionDocsPendientes.rpt")
      Case "Recepcion"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecepcionDocsTags.rpt")
      Case "Devolucion"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecepcionDocsTags.rpt")
    End Select

 .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy  hh:mm:ss") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .StoredProcParam(0) = Format(dtpFInicio.Value, "yyyy-MM-dd 00:00:00.000")
 .StoredProcParam(1) = Format(dtpFFin.Value, "yyyy-MM-dd 23:59:59.000")
 
 Select Case pTipo
  Case "Inventario"
     .StoredProcParam(2) = mTagRecepcion
     .StoredProcParam(3) = mTagDevolucion
  Case "Pendientes"
    If optPendRecepcion.Value = True Then
        .Formulas(2) = "fxTitulo='Pendientes de Recepción de Documentación'"
        .StoredProcParam(2) = "N"
    Else
        .Formulas(2) = "fxTitulo='Pendientes por Devolución de Documentación'"
        .StoredProcParam(2) = "D"
    End If
  Case "Recepcion"
     .StoredProcParam(2) = mTagRecepcion
     .StoredProcParam(3) = cboUsuario.Text
  Case "Devolucion"
     .StoredProcParam(2) = mTagDevolucion
     .StoredProcParam(3) = cboUsuario.Text
End Select

 .PrintReport

End With

Me.MousePointer = vbDefault


End Sub


Private Sub sbCargarGridConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    If txtOperacionBuscar.Text = Empty Then Exit Sub

    Me.MousePointer = vbHourglass

    strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO from CRD_OPERACION_TAGS OT" _
           & " inner join CRD_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO where OT.ID_SOLICITUD = " & txtOperacionBuscar _
           & " and (T.TAG_CODIGO = '" & mTagRecepcion & "' or T.TAG_CODIGO = '" & mTagDevolucion & "')"
            
    vGridConsulta.MaxCols = 3
    vGridConsulta.MaxRows = 0


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridConsulta.MaxRows = vGridConsulta.MaxRows + 1
    vGridConsulta.Row = vGridConsulta.MaxRows
  
    vGridConsulta.Col = 1
    vGridConsulta.Text = rs!Descripcion
    
    vGridConsulta.Col = 2
    vGridConsulta.Value = IIf(IsNull(rs!REGISTRO_FECHA), "", rs!REGISTRO_FECHA)
    
    vGridConsulta.Col = 3
    vGridConsulta.Value = IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO)
    
    vGridConsulta.RowHeight(vGridConsulta.Row) = vGridConsulta.MaxTextRowHeight(vGridConsulta.Row)
    rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCargarUsuarios()
Dim strSQL As String

On Error GoTo vError
    Me.MousePointer = vbHourglass

    strSQL = "SELECT UPPER(NOMBRE) as ItmX from USUARIOS WHERE ESTADO = 'A'"
    
    Call sbLlenaCbo(cboUsuario, strSQL, True)

    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtOperacionBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub
