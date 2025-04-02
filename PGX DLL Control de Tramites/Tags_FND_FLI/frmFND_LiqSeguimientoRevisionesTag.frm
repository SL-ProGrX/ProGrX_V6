VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFNDLiqSeguimientoRevisionesTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Liquidaciones Fondos"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8280
      Top             =   120
   End
   Begin VB.Frame FraControles 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   11175
      Begin TabDlg.SSTab SSTab 
         Height          =   6615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11668
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Liquidaciones"
         TabPicture(0)   =   "frmFND_LiqSeguimientoRevisionesTag.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "imgRefresh"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "tlbRefresh"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vGrid"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkTipoSalidaLiq"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Seguimiento"
         TabPicture(1)   =   "frmFND_LiqSeguimientoRevisionesTag.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vGridSeguimiento"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Revisión"
         TabPicture(2)   =   "frmFND_LiqSeguimientoRevisionesTag.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtObservacion"
         Tab(2).Control(1)=   "cboEtiquetas"
         Tab(2).Control(2)=   "tlbAplicar"
         Tab(2).Control(3)=   "lswErrores"
         Tab(2).Control(4)=   "Label8(1)"
         Tab(2).Control(5)=   "Label2(0)"
         Tab(2).Control(6)=   "Label27"
         Tab(2).ControlCount=   7
         Begin VB.CheckBox chkTipoSalidaLiq 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Mostrar Liquidaciones con Desembolso"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6480
            TabIndex        =   19
            Top             =   480
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.TextBox txtObservacion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   -73440
            MaxLength       =   995
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   4020
            Width           =   9135
         End
         Begin VB.ComboBox cboEtiquetas 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmFND_LiqSeguimientoRevisionesTag.frx":68A6
            Left            =   -73440
            List            =   "frmFND_LiqSeguimientoRevisionesTag.frx":68A8
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   540
            Width           =   5295
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   5415
            Left            =   120
            TabIndex        =   6
            Top             =   1020
            Width           =   10695
            _Version        =   524288
            _ExtentX        =   18865
            _ExtentY        =   9551
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
            MaxCols         =   10
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmFND_LiqSeguimientoRevisionesTag.frx":68AA
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridSeguimiento 
            Height          =   5775
            Left            =   -74880
            TabIndex        =   8
            Top             =   660
            Width           =   10575
            _Version        =   524288
            _ExtentX        =   18653
            _ExtentY        =   10186
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
            MaxCols         =   487
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmFND_LiqSeguimientoRevisionesTag.frx":7472
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin MSComctlLib.Toolbar tlbAplicar 
            Height          =   570
            Left            =   -65520
            TabIndex        =   11
            Top             =   5820
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1005
            ButtonWidth     =   2117
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "Aplicar"
                  Object.ToolTipText     =   "Aplicar Etiqueta"
                  ImageKey        =   "IMG1"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lswErrores 
            Height          =   2655
            Left            =   -73440
            TabIndex        =   12
            Top             =   1140
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Aplicado"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Mensaje"
               Object.Width           =   8819
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbRefresh 
            Height          =   330
            Left            =   9600
            TabIndex        =   18
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            ButtonWidth     =   1984
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgRefresh"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Refrescar"
                  Key             =   "Refrescar"
                  Object.ToolTipText     =   "Volver a cargar la información"
                  ImageIndex      =   1
               EndProperty
            EndProperty
            MousePointer    =   1
         End
         Begin MSComctlLib.ImageList imgRefresh 
            Left            =   8880
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":7A8C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -74880
            TabIndex        =   15
            Top             =   4020
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Etiqueta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   -74880
            TabIndex        =   14
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Omisiones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74880
            TabIndex        =   13
            Top             =   1140
            Width           =   855
         End
      End
   End
   Begin VB.Frame fraOperacion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      Begin VB.TextBox txtContrato 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtPlan 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtCedula 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label LblOperacion 
         Caption         =   "Cedula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Nombre"
         Top             =   120
         Width           =   5295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":7BB1
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":E413
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":14C75
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":1B4D7
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":21D39
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2859B
            Key             =   "IMG6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   9360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2EDFD
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2EF1B
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F041
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F16B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F27D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F394
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F495
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F5CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F6E1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":2F805
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":36067
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":3C8C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":3C9E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNombreUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1920
      Picture         =   "frmFND_LiqSeguimientoRevisionesTag.frx":3CB01
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmFNDLiqSeguimientoRevisionesTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCedula As String, mConsecutivo As String, mLlave As String
Dim mPlan As String, mContrato As Long, vPaso As Boolean

Private Sub sbCargarObservacion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from SIF_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtObservacion = rs.Fields(0) & vbNewLine
    Else
        txtObservacion = Empty
    End If
    
    For i = 1 To lswErrores.ListItems.Count
        If lswErrores.ListItems(i).Checked = True Then
            If lswErrores.ListItems(i).SubItems(2) = "N" Then
                If txtObservacion = Empty Then
                    txtObservacion.Text = "-" & lswErrores.ListItems(i).SubItems(3)
                Else
                    txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & lswErrores.ListItems(i).SubItems(3)
                End If
            End If
        End If
    Next
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cboEtiquetas_Click()
If vPaso Then Exit Sub
Call sbCargarObservacion
End Sub


Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
vModulo = 8


lblNombreUsuario.Caption = glogon.Usuario

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lswErrores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
              If txtObservacion = Empty Then
                txtObservacion.Text = " - " & Item.SubItems(1)
              Else
                txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
              End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert SIF_OMISIONESG (cedula,ID_ERROR,MODULO,CODIGO,DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO) values('" & mCedula _
             & "'," & Item.Text & ",'FLQ','" & mLlave & "','" & mConsecutivo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
      Call ConectionExecute(strSQL)
             
      strSQL = "select max(LINEA_ERR) as 'Linea' from SIF_OMISIONESG where codigo = '" & mLlave & "' and Documento = '" & mConsecutivo & "' and ID_ERROR = " & Item.Text
      Call OpenRecordSet(rs, strSQL)
          Item.Tag = rs!Linea
      rs.Close
      
      If txtObservacion = Empty Then
        txtObservacion.Text = " - " & Item.SubItems(1)
      Else
        txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
      End If
      
    Else
      strSQL = "delete SIF_OMISIONESG where LINEA_ERR = " & Item.Tag
      Call ConectionExecute(strSQL)
      Item.Tag = ""

      Call sbCargarObservacion
    End If
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Select Case SSTab.Tab
   Case 1
     Call sbCargarGridSeguimiento(txtCedula)
   Case 2
     Call sbCargarListaErrores
     Call sbCargarCombosEtiquetas
End Select
End Sub



Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargarListaLiquidaciones

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError

Me.MousePointer = vbHourglass

If Trim(cboEtiquetas.Text) = Empty Then
    MsgBox "Debe seleccionar la etiqueta que desea plicar"
    Me.MousePointer = vbDefault
    Exit Sub
End If

If MsgBox("Está seguro que sea aplicar la etiqueta en La liquidacion de Fondos seleccionada", vbExclamation + vbYesNo) = vbNo Then
    Me.MousePointer = vbDefault
    Exit Sub
End If

Call sbSIFRegistraTags(mCedula, SIFGlobal.fxCodText(cboEtiquetas.Text), txtObservacion, mConsecutivo, "FLQ", mConsecutivo)


Call sbAplicarErrores
Call sbCargarListaLiquidaciones

txtCedula.SetFocus
SSTab.Tab = 0


Me.MousePointer = vbDefault
Exit Sub

vError:
Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbRefresh_ButtonClick(ByVal Button As MSComctlLib.Button)
Call TimerX_Timer
End Sub

Private Sub txtCedula_GotFocus()
SSTab.Tab = 0
mCedula = Empty
mConsecutivo = Empty

lblNombre = Empty
SSTab.TabEnabled(1) = False
SSTab.TabEnabled(2) = False


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtCedula.Text) <> "" Then
  Call sbCargarListaLiquidaciones(txtCedula)
  lblNombre.Caption = fxNombre(txtCedula.Text)
End If
End Sub


Private Sub txtCedula_LostFocus()
If Trim(txtCedula.Text) <> "" Then
   Call sbCargarListaLiquidaciones(txtCedula.Text)
   lblNombre.Caption = fxNombre(txtCedula.Text)
End If
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    vGrid.Row = Row
    vGrid.Col = 2
    mConsecutivo = vGrid.Text

    vGrid.Col = 3
    mCedula = vGrid.Text
    txtCedula.Text = mCedula
    
    vGrid.Col = 4
    lblNombre.Caption = vGrid.Text
    vGrid.Col = 7
    mPlan = vGrid.Text
    vGrid.Col = 8
    mContrato = vGrid.Text
    
   
    txtPlan.Text = mPlan
    txtContrato.Text = mContrato
    txtContrato.ToolTipText = "No. Liquidación..: " & mConsecutivo
    
    mLlave = Trim(mPlan) & "-" & mContrato
    
    If Len(Trim(mCedula)) > 0 Then
      SSTab.TabEnabled(1) = True
     SSTab.TabEnabled(2) = True
    End If

End Sub


Private Sub sbCargarListaLiquidaciones(Optional ByVal strCedula As String = Empty)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "SELECT Top 3000 'B',  L.CONSEC, F.CEDULA,S.nombre,L.usuario,L.FECHA,L.cod_plan,L.cod_Contrato" _
       & ", case when L.Retencion_Codigo is null then 'No' else 'Sí' end as 'Retiene'" _
       & ", Ban.Descripcion " _
       & " FROM fnd_liquidacion L  inner join FND_CONTRATOS F" _
       & " on L.COD_PLAN = F.COD_PLAN and L.COD_CONTRATO = F.COD_CONTRATO AND L.COD_OPERADORA = F.COD_OPERADORA" _
       & " inner join SOCIOS S on F.CEDULA = S.CEDULA" _
       & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
       & " LEFT JOIN TES_BANCOS Ban on L.cod_Banco = Ban.Id_Banco" _
       & " WHERE  isnull(L.ANALISTA_REVISION,'N') = 'N'"
           
If Trim(txtCedula.Text) <> "" Then strSQL = strSQL & " and F.Cedula = '" & txtCedula.Text & "'"
  
If chkTipoSalidaLiq.Value = vbChecked Then
   strSQL = strSQL & " and L.Retencion_Codigo is null"
End If
           
vPaso = True
Call sbCargaGrid(vGrid, 10, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = False

Me.MousePointer = vbDefault

Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarGridSeguimiento(ByVal mCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
If mContrato = Empty Then Exit Sub

Me.MousePointer = vbHourglass

vGridSeguimiento.MaxCols = 4
vGridSeguimiento.MaxRows = 0


strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO" _
       & " from SIF_CONTROL_TAGS OT inner join SIF_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO" _
       & " where OT.documento = '" & mConsecutivo & "' and cod_Modulo = 'FLQ'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!REGISTRO_USUARIO & "[" & rs!REGISTRO_FECHA & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!notas), "", rs!notas)
    
    vGridSeguimiento.Col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!REGISTRO_FECHA), "", rs!REGISTRO_FECHA)
    
    vGridSeguimiento.Col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
    rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCargarCombosEtiquetas()
Dim strSQL As String

On Error GoTo vError

    
    strSQL = "SELECT CT.TAG_CODIGO + ' - ' +  rtrim(CT.DESCRIPCION) as 'ItmX'" _
            & " FROM SIF_TAGS CT INNER JOIN SIF_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
            & " INNER JOIN SIF_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
            & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
            & "' and  CT.TAG_CODIGO in(select TAG_CODIGO from SIF_TAGS_MODULOS where cod_modulo = 'FLQ')" _
            & " order by CT.TAG_CODIGO"
    vPaso = True
    Call sbLlenaCbo(cboEtiquetas, strSQL, False, False)
    vPaso = False
    Call cboEtiquetas_Click
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargarListaErrores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If txtCedula = Empty Then
    Exit Sub
End If

With lswErrores
 .ListItems.Clear
 
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE, ER.LINEA_ERR" _
        & " from sif_Omisiones E left join SIF_OMISIONESG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.cedula = '" & txtCedula.Text & "' and ER.Modulo = 'FLQ' and ER.Codigo = '" & mLlave _
        & "' and ER.Documento = '" & mConsecutivo & "'" _
        & " where E.ACTIVO = '1'  and E.ID_ERROR in(select ID_ERROR from SIF_OMISIONES_MODULOS where cod_modulo = 'FLQ') " _
        & " order by E.ID_ERROR"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
         itmX.Tag = rs!LINEA_ERR
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub


Private Sub sbAplicarErrores()
'' Procedimiento para colocar los errores ingresados en aplicados
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    strSQL = "update SIF_OMISIONESG SET APLICADO = 'S' WHERE cedula = '" & mCedula _
           & "' AND MODULO = 'FLQ' AND CODIGO = '" & mLlave & "' AND DOCUMENTO = '" & mConsecutivo & "'"
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

