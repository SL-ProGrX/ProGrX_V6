VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAF_CD_Liquidaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidaciones de Desembolsos por Cómites Sedes"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14895
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11880
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
      TabCaption(0)   =   "Desembolsos"
      TabPicture(0)   =   "frmAF_CD_Liquidaciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tlbDetallarLiquidacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detallar Liquidación"
      TabPicture(1)   =   "frmAF_CD_Liquidaciones.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Historico"
      TabPicture(2)   =   "frmAF_CD_Liquidaciones.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vGridHistorico"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Detalle Liquidación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -70680
         TabIndex        =   13
         Top             =   4560
         Width           =   10215
         Begin VB.TextBox txtNotas 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   720
            Width           =   8655
         End
         Begin VB.TextBox txtTotalFactura 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   15
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtDiferencia 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7920
            MaxLength       =   15
            TabIndex        =   14
            Top             =   240
            Width           =   2055
         End
         Begin MSComctlLib.Toolbar tlbLiquidar 
            Height          =   360
            Left            =   8880
            TabIndex        =   17
            Top             =   1440
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   635
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "Liquidacion"
                  ImageIndex      =   4
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Notas de la liquidación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   10080
            X2              =   120
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label5 
            Caption         =   "Monto Documentos"
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
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Total a Liquidar"
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
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Diferencia"
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
            Left            =   7080
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuentas"
         Height          =   6015
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   4095
         Begin FPSpreadADO.fpSpread vGridOpxDetallar 
            Height          =   5580
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3855
            _Version        =   524288
            _ExtentX        =   6800
            _ExtentY        =   9843
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":0054
            VScrollSpecialType=   2
            Appearance      =   1
            AppearanceStyle =   1
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Detalle Facturas"
         Height          =   3975
         Left            =   -70680
         TabIndex        =   9
         Top             =   480
         Width           =   10215
         Begin FPSpreadADO.fpSpread vGridFacturas 
            Height          =   3615
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   9975
            _Version        =   524288
            _ExtentX        =   17595
            _ExtentY        =   6376
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
            MaxCols         =   5
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":0679
            VScrollSpecialType=   2
            Appearance      =   1
            AppearanceStyle =   1
         End
      End
      Begin FPSpreadADO.fpSpread vGridHistorico 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   14055
         _Version        =   524288
         _ExtentX        =   24791
         _ExtentY        =   10398
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
         MaxCols         =   15
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":0CED
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbDetallarLiquidacion 
         Height          =   360
         Left            =   10920
         TabIndex        =   7
         Top             =   6120
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   635
         ButtonWidth     =   3493
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Detallar Liquidación"
               Key             =   "Detallar"
               ImageIndex      =   7
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5415
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   13815
         _Version        =   524288
         _ExtentX        =   24368
         _ExtentY        =   9551
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
         MaxCols         =   10
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":1D2D
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   1296
      BandCount       =   4
      _CBWidth        =   14895
      _CBHeight       =   735
      _Version        =   "6.7.9782"
      Child1          =   "txtCodigoComite"
      MinHeight1      =   315
      Width1          =   2265
      NewRow1         =   0   'False
      Child2          =   "txtDescripcionComite"
      MinHeight2      =   315
      Width2          =   5640
      NewRow2         =   0   'False
      Child3          =   "txtRate"
      MinWidth3       =   165
      MinHeight3      =   315
      Width3          =   165
      NewRow3         =   0   'False
      Child4          =   "tlbMenu"
      MinHeight4      =   330
      Width4          =   4095
      NewRow4         =   -1  'True
      Begin VB.TextBox txtCodigoComite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   30
         Width           =   2070
      End
      Begin MSComctlLib.Toolbar tlbMenu 
         Height          =   330
         Left            =   165
         TabIndex        =   4
         Top             =   375
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos"
               Key             =   "Datos"
               Object.ToolTipText     =   "Actualiza Datos Personales"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cuentas"
               Key             =   "Cuentas"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reportes"
               Key             =   "Reportes"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDescripcionComite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   5445
      End
      Begin VB.TextBox txtRate 
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
         Height          =   315
         Left            =   8130
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Exposición a Riesgo de la persona"
         Top             =   30
         Width           =   6675
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   360
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
            Picture         =   "frmAF_CD_Liquidaciones.frx":258A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":176FC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":17FD6
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":2D148
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":422BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":58C7C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":59556
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":6FF18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":733AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Liquidaciones.frx":7683C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAF_CD_Liquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim vOperacion As String, vDocumento As String, vDeposito As String, vDetalle As String, vFecha As String
Dim vMonto As Double

Private Sub Form_Activate()
 vModulo = 23
End Sub

Private Sub Form_Load()
  vModulo = 23
  txtTotal.Text = 0
  txtTotal.Text = 0
  txtTotalFactura.Text = 0
  txtDiferencia.Text = 0
  SSTab.Tab = 0
  vGrid.MaxRows = 0
  vGridOpxDetallar.MaxRows = 0
  vGridFacturas.MaxRows = 0
  vGridFacturas.MaxCols = 5
  vGridHistorico.MaxRows = 0
  If GLOBALES.gTag <> Empty Then txtCodigoComite.Text = GLOBALES.gTag
  GLOBALES.gTag = ""
End Sub

Private Sub sbGuardaFactura()
  
  strSQL = "INSERT AFI_CD_DETALLE_LIQUIDACION (NOPERACION, NDOCUMENTO,DEPOSITO, DETALLE, " _
         & "FECHA_DOCUMENTO, MONTO,REGISTRO_FECHA, REGISTRO_USUARIO) Values " _
         & "(" & vOperacion & ",'" & vDocumento & "','" & vDeposito & "','" & vDetalle & "' " _
         & ", '" & Format(vFecha, "yyyymmdd") & "'," & vMonto & ",getdate(),'" & glogon.Usuario & "')"
  glogon.Conection.Execute strSQL
   
End Sub

Private Sub sbModificaFactura()

  strSQL = "UPDATE AFI_CD_DETALLE_LIQUIDACION SET DETALLE ='" & vDetalle & "',FECHA_DOCUMENTO ='" & Format(vFecha, "yyyymmdd") & "'" _
         & ", MONTO=" & vMonto & ",REGISTRO_FECHA =getdate(),REGISTRO_USUARIO ='" & glogon.Usuario & "'" _
         & " WHERE NOPERACION=" & vOperacion & " and NDOCUMENTO ='" & vDocumento & "' "
  glogon.Conection.Execute strSQL

End Sub

Private Sub sbEliminaFactura()
  
  strSQL = "DELETE FROM AFI_CD_DETALLE_LIQUIDACION WHERE NOPERACION=" & vOperacion & " and NDOCUMENTO ='" & vDocumento & "' "
  glogon.Conection.Execute strSQL
  
End Sub
'Trae los Montos de la liquidación
'y el monto total en facturas registradas
Private Sub sbTraeMontos(ByVal vOperacion As Integer)
Dim Saldo, MontoTotal, MontoFacturas As Double
Dim TesoreriaSolucitud As Long
On Error GoTo vError

txtTotal.Text = 0
txtTotalFactura.Text = 0
txtDiferencia.Text = 0

'Se obtiene el Monto Total de la Operacion
strSQL = "Select Monto,NOPERACION from AFI_CD_CUENTAS " _
       & " Where Noperacion = " & vOperacion & ""
rs.Open strSQL, glogon.Conection, adOpenStatic

MontoTotal = Format(rs!Monto, "standard")

rs.Close

MontoFacturas = 0
'Se obtiene el total de las facturas que respaldan la Liquidación
strSQL = "Select isnull(MONTO,0) as 'Monto' from AFI_CD_DETALLE_LIQUIDACION " _
       & " Where NOPERACION = " & vOperacion & " "
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  MontoFacturas = MontoFacturas + Format(rs!Monto, "standard")
  rs.MoveNext
Loop


rs.Close
          
Saldo = MontoTotal - MontoFacturas
  
txtTotal.Text = Format(CDbl(txtTotal.Text) + MontoTotal, "standard")
txtTotalFactura.Text = Format(MontoFacturas, "standard")
txtDiferencia.Text = Format(Saldo, "standard")

Exit Sub
vError:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub sbActualizaSaldo(ByVal vOperacion As Integer)
Dim TotalSaldo As Double
On Error GoTo vError

strSQL = "Select isnull(C.Monto - sum(DL.MONTO),0) as 'Saldo'" _
       & " from dbo.AFI_CD_CUENTAS C" _
       & "  inner join AFI_CD_DETALLE_LIQUIDACION DL on C.NOPERACION = DL.NOPERACION" _
       & " Where C.Noperacion = " & vOperacion
          
rs.Open strSQL, glogon.Conection, adOpenStatic

TotalSaldo = Format(rs!Saldo, "Standard")

rs.Close

strSQL = "Update AFI_CD_CUENTAS set SALDO = " & CCur(TotalSaldo) & ",ESTADO ='L'" _
       & " where NOPERACION= " & vOperacion & ""
glogon.Conection.Execute strSQL

Exit Sub
vError:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case SSTab.Tab
      Case 1
        Call sbCargaOpxDetallar
      Case 2
        Call sbCargaHistorico
     End Select
End Sub

Private Sub tlbDetallarLiquidacion_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError

Dim vTotal As Double
Dim strSQL As String
Dim i As Integer

    If Trim(txtTotal.Text) = "" Then txtTotal.Text = 0
    
    vTotal = CDbl(txtTotal.Text)
    
    With vGrid
     For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
         .Col = 2
         vOperacion = .Text
            
         'D= Detalle
         strSQL = "update afi_cd_cuentas set PROCESO = 'D' " _
                & "where NOPERACION = '" & vOperacion & "'"
         glogon.Conection.Execute strSQL
            
        End If
        
     Next i
    End With
    Call sbCuentaOpPendientes
    Call sbCargaOperaciones
    Call sbCargaOpxDetallar
    
    txtTotal.Text = Format(vTotal, "Standard")
    
Exit Sub
vError:
  MsgBox Err.Description

End Sub

Private Sub tlbLiquidar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
Dim vTipoDoc As String, vTransaccion As String
 
If vOperacion = "" Or vOperacion = "0" Then Exit Sub
 
If CCur(txtDiferencia.Text) > 0 Then
   MsgBox "Existen Diferencias en el detalle con el monto a Cancelar...Revise!", vbExclamation
   Exit Sub
End If
 
With vGridOpxDetallar
  For i = 1 To .MaxRows
   .Row = i
   .Col = 1
   If .Value = vbChecked Then
      .Col = 2
      vOperacion = .Text
      strSQL = "exec spAFI_CD_AsientoLiquidacion " & vOperacion & ",'" & glogon.Usuario & "','" _
             & GLOBALES.gOficinaTitular & "','" & txtNotas.Text & "'"
      
      rs.Open strSQL, glogon.Conection, adOpenStatic
        vTipoDoc = rs!TipoDoc
        vTransaccion = rs!transaccion
      rs.Close
   End If
Next i

If GLOBALES.SysDocVersion = 2 Then
  Call sbImprimeRecibo(vTransaccion, vTipoDoc)
End If

MsgBox "La Liquidación fue registrada satisfactoriamente", vbInformation, "Información"

Call sbCargaOpxDetallar

End With
    
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError

GLOBALES.gTag = txtCodigoComite.Text
Select Case Button.Key
  Case "Datos"
    Call sbSIFForms("frmAF_CD_Comites", , , , False, Me)
        
  Case "Cuentas"
    Call sbSIFForms("frmAF_CD_Cuentas", , , , False, Me)
    
  Case "Reportes"
      strSQL = ""
      With frmContenedor.Crt
         .Reset
         .WindowShowGroupTree = True
         .WindowShowPrintSetupBtn = True
         .WindowShowRefreshBtn = True
         .WindowShowSearchBtn = True
         .WindowState = crptMaximized
         .Connect = glogon.ConectRPT
         .WindowTitle = "Reporte consulta de movimiento de actividades"
         .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ControlLiquidacionEspecifico.rpt")
         .Formulas(0) = "fxTitulo= 'CONTROL DE LIQUIDACIONES POR COMITE'"
         strSQL = "({afi_cd_cuentas.cod_comite}) = '" & txtCodigoComite.Text & "'"
         .SelectionFormula = strSQL
         .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
         .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
         .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
         
         .PrintReport
      End With
End Select

Exit Sub
vError:
      MsgBox Err.Description, vbCritical

End Sub
Private Sub txtCodigoComite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      Call sbCargaComites
      Call sbCuentaOpPendientes
      Call sbCargaOperaciones
    ElseIf KeyCode = vbKeyF4 Then
        gBusquedas.Columna = "C.COD_COMITE"
        gBusquedas.Orden = "C.COD_COMITE"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select C.COD_COMITE,CM.DESCRIPCION" _
                            & " from AFI_CD_CUENTAS C" _
                            & " inner join AFI_CD_COMITES CM on c.COD_COMITE = CM.COD_COMITE and C.ESTADO ='T' "
        frmBusquedas.Show vbModal
        txtCodigoComite = gBusquedas.Resultado
        txtDescripcionComite = gBusquedas.Resultado2
        
        Call sbCuentaOpPendientes
        Call sbCargaOperaciones
    End If
     SSTab.Tab = 0
     GLOBALES.gTag = Trim(txtCodigoComite.Text)
End Sub

Private Sub sbCuentaOpPendientes()
 Dim strSQL As String, rs As New ADODB.Recordset
    strSQL = "select count(COD_COMITE)as Cuenta " _
           & "from AFI_CD_CUENTAS where estado='T' and COD_COMITE='" & Trim(txtCodigoComite.Text) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    txtRate.Text = "Liquidaciones Pendientes : " & IIf(IsNull(rs!Cuenta), 0, rs!Cuenta)
     
    rs.Close
 
End Sub

Private Sub sbCargaComites()
 Dim strSQL As String, rs As New ADODB.Recordset
 
  strSQL = "select DESCRIPCION from AFI_CD_COMITES where COD_COMITE='" & Trim(txtCodigoComite.Text) & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
      
  If Not rs.EOF Then
    txtDescripcionComite.Text = Trim(rs!Descripcion)
  End If
  
  rs.Close
End Sub

Private Sub sbCargaOpxDetallar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

vGridFacturas.MaxRows = 0
vGridFacturas.MaxRows = 1


  With vGridOpxDetallar
   .MaxRows = 1
   
   strSQL = "select C.NOPERACION as Operacion,datediff(d,C.REGISTRO_FECHA,getdate()) as 'Dias_Pendientes' " _
         & ", sum(CA.MONTO) as Monto from dbo.AFI_CD_CUENTAS C" _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " where  C.COD_COMITE = '" & txtCodigoComite.Text & "' and C.Estado = 'T' and C.PROCESO ='D'" _
         & " group by C.NOPERACION,C.REGISTRO_FECHA"
         

  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  Do While Not rs.EOF
    .Row = .MaxRows
    
    .Col = 2
    .Text = rs!Operacion
    
    .Col = 3
    .Text = Format(rs!Monto, "Standard")
    
    .Col = 4
    .Text = rs!Dias_Pendientes
                
    .MaxRows = .MaxRows + 1
    rs.MoveNext
    
  Loop
  .MaxRows = .MaxRows - 1
  rs.Close
  
 End With

Exit Sub

vError:
      MsgBox Err.Description, vbCritical

End Sub

Private Sub sbCargaOperaciones()
On Error GoTo error
Dim strSQL As String, rs As New ADODB.Recordset
Dim vItem As MSComctlLib.ListItem
Dim vLvw As MSComctlLib.ListView
Dim vKey As String
Dim i As Integer

  
   strSQL = "select C.NOPERACION,C.ACTIVA_FECHA , DATEDIFF(D,C.ACTIVA_FECHA,GETDATE()) as 'Dias_Pendientes' " _
         & ",CA.MONTO as Monto,A.DESCRIPCION ACTIVIDAD,case C.ESTADO when 'T'  then 'Trasladado' when 'A'  then 'Activo' " _
         & "Else 'Liquidado' End as Estado,case C.TIPO when 'T' then 'Transferencia' else 'Cheque' End as Desembolso " _
         & ",C.REGISTRO_USUARIO, Tes.FECHA_EMISION" _
         & " from dbo.AFI_CD_CUENTAS C" _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " inner join AFI_CD_ACTIVIDADES A on CA.COD_ACTIVIDAD = A.COD_ACTIVIDAD" _
         & "  left join TES_TRANSACCIONES Tes on C.TESORERIA_NSOLICITUD = Tes.NSOLICITUD" _
         & " where  C.COD_COMITE='" & txtCodigoComite.Text & "' and C.Estado='T' and C.PROCESO='T'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  With vGrid
    .MaxRows = 1
    
    For i = 1 To .MaxCols
     .Col = i
     .Text = ""
    Next i
          
    Do While Not rs.EOF
      .Row = .MaxRows
      
      .Col = 2
      .Text = IIf(IsNull(rs!Noperacion), 0, rs!Noperacion)
                
      .Col = 3
      .Text = rs!ACTIVA_FECHA & ""
      
      .Col = 4
      .Text = rs!FECHA_EMISION & ""
      
      .Col = 5
      .Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "standard")
      
      .Col = 6
      .Text = rs!ACTIVIDAD
      
      .Col = 7
      .Text = rs!Dias_Pendientes
      
      .Col = 8
      .Text = rs!Estado
      
      .Col = 9
      .Text = rs!desembolso
      
      .Col = 10
      .Text = rs!REGISTRO_USUARIO
      
      rs.MoveNext
     .MaxRows = .MaxRows + 1
    
    Loop
    
    rs.Close
   .MaxRows = .MaxRows - 1
     
  End With

Exit Sub

error:
      MsgBox Err.Description

End Sub

Public Function fxGridSumaFacturas(vGrid As Object, Columna As Long) As Double
' Este procedimiento valida que solo pueda haber una registro marcado en el grid
Dim suma As Double, i As Long
    

On Error GoTo vError

    suma = 0
    vGrid.Row = 1
    vGrid.Col = 1
    For i = 1 To vGrid.MaxRows
      vGrid.Row = i
      vGrid.Col = Columna
      If IsNumeric(vGrid.Value) Then
         suma = suma + vGrid.Value
      End If
         vGrid.Col = 1
    Next i
    fxGridSumaFacturas = suma
    
Exit Function

vError:
        MsgBox Err.Description
    
End Function

'Modifica las facturas y trae el monto
Private Sub vGridFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

With vGridFacturas
  .Row = .ActiveRow
  
  If .ActiveCol = .MaxCols And KeyCode = vbKeyReturn Then
    .Col = 1
    If Not fxValidaFactura Then
       Call sbGuardaFactura
       .MaxRows = .MaxRows + 1
    Else
       Call sbModificaFactura
       Call sbTraeFacturas
    End If

  End If
  
 If .ActiveCol = .MaxCols And KeyCode = vbKeyDelete Then
     If fxValidaFactura Then
        Call sbEliminaFactura
        Call sbTraeFacturas
     End If
 End If

End With
  
Call sbTraeMontos(vOperacion)
  
Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Function fxValidaFactura() As Boolean
On Error GoTo error

 If vOperacion = Empty Then
    MsgBox "Debe seleccionar una operación"
    Exit Function
 End If
 
 
 With vGridFacturas
    
    .Row = .ActiveRow
    
    .Col = 1
    vDeposito = .Value
    
    .Col = 2
    If .Text = Empty Then
      MsgBox "Falta número de documento"
    Else
      vDocumento = .Text
    End If
    
    .Col = 3
    If .Text = Empty Then
      MsgBox "Falta Fecha de documento"
    Else
      vFecha = .Text
    End If
    
    .Col = 4
    If .Text = Empty Then
      MsgBox "Falta detalle de documento"
    Else
      vDetalle = .Text
    End If
    
    .Col = 5
    If .Text = Empty Then
      MsgBox "Falta monto de documento"
    Else
      vMonto = CCur(.Text)
      .Text = Format(.Text, "standard")
    End If
    
    strSQL = "SELECT NOPERACION,NDOCUMENTO" _
           & " FROM AFI_CD_DETALLE_LIQUIDACION" _
           & " where NOPERACION = " & vOperacion & " and NDOCUMENTO='" & Trim(vDocumento) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    If rs.EOF = True Then
      fxValidaFactura = False
    ElseIf rs!Noperacion = vOperacion And vDocumento = Trim(vDocumento) Then
       fxValidaFactura = True
    Else
       
    End If

    rs.Close
    
 End With


Exit Function
error:
     MsgBox Err.Description, vbCritical
     
End Function

Private Sub sbCargaHistorico()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError



  strSQL = "select C.NOPERACION as Operacion,C.NOTAS,C.LIQUIDA_FECHA,C.ACTIVA_FECHA,Tes.FECHA_EMISION" _
         & ", CA.MONTO as Monto,A.DESCRIPCION ACTIVIDAD,C.TESORERIA_FECHA,C.TESORERIA_NSOLICITUD" _
         & ", case C.ESTADO when 'T'  then 'Trasladado' when 'A'  then 'Activo' " _
         & " Else 'Liquidado' End as Estado ,Case C.APRUEBA when 'J' then 'junta Directiva' when 'O' then 'Oficina Regional' Else" _
         & "' Director Zona' End as Aprueba,case C.TIPO when 'T' then 'Transferencia' else 'Cheque' End as Desembolso,C.REGISTRO_FECHA,C.REGISTRO_USUARIO" _
         & ", Tes.Beneficiario as 'TESORERIA_BENEFICIARIO', Tes.Codigo AS 'TESORERIA_CODIGO'" _
         & " from dbo.AFI_CD_CUENTAS C " _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " inner join AFI_CD_ACTIVIDADES A on CA.COD_ACTIVIDAD = A.COD_ACTIVIDAD" _
         & " left join TES_Transacciones Tes on C.TESORERIA_NSOLICITUD = Tes.Nsolicitud" _
         & " where  C.COD_COMITE='" & txtCodigoComite.Text & "' order by C.REGISTRO_FECHA desc"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  With vGridHistorico
    .MaxRows = 1
    .Row = .MaxRows
         
    Do While Not rs.EOF
      .Row = .MaxRows
      
      .Col = 3
      .Text = IIf(IsNull(rs!Operacion), 0, rs!Operacion)
      
      .Col = 4
      .Text = rs!ACTIVIDAD
      
      .Col = 5
      .Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "standard")
      
      .Col = 6
      .Text = rs!Estado
      
      .Col = 7
      .Text = rs!TESORERIA_FECHA & ""
      
      .Col = 8
      .Text = rs!REGISTRO_USUARIO & ""
      
      .Col = 9 'Fecha Activacion
      .Text = rs!ACTIVA_FECHA & ""
      
      .Col = 10 'Fecha Liquidacion
      .Text = rs!LIQUIDA_FECHA & ""
      
      .Col = 11 'No. Solicitud Tesoreria
      .Text = rs!TESORERIA_NSOLICITUD & ""
      
      .Col = 12 'Fecha de Pago Real en Tesorería
      .Text = rs!FECHA_EMISION & ""
      
      .Col = 13 'Beneficiario Tesoreria
      .Text = rs!TESORERIA_CODIGO & ""
      
      .Col = 14 'Beneficiario Tesoreria
      .Text = rs!TESORERIA_BENEFICIARIO & ""
      
      .Col = 15 'notas
      .Text = rs!NOTAS & ""
      
      rs.MoveNext
     .MaxRows = .MaxRows + 1
    Loop
   .MaxRows = .MaxRows - 1
     
  End With

  rs.Close
  
Exit Sub

vError:
      MsgBox Err.Description, vbCritical

End Sub



Private Sub vGridHistorico_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String
Dim vOpRef As String

With vGridHistorico
  .Row = Row
  .Col = 3
  vOpRef = .Text
  vNumDoc = vOpRef
  
  If Col = 1 Then
     vTipoDoc = "CD.CxC"
  Else
     vTipoDoc = "CD.Liq"
     strSQL = "select cod_Transaccion from sif_transacciones" _
            & " where Tipo_Documento = '" & vTipoDoc & "' and Referencia_01 = '" & txtCodigoComite.Text _
            & "' and Referencia_02 = '" & vOpRef & "'"
     rs.Open strSQL, glogon.Conection, adOpenStatic
     If Not rs.EOF And Not rs.BOF Then
         vNumDoc = rs!Cod_Transaccion
     End If
     rs.Close
  End If
  
End With

Call sbImprimeRecibo(vNumDoc, vTipoDoc)

End Sub

Private Sub vGridOpxDetallar_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo error
Dim i As Integer
Dim vTotal As Double

txtTotal.Text = 0
txtDiferencia.Text = 0
vTotal = 0

With vGridOpxDetallar
   For i = 1 To .MaxRows
     .Row = i
     .Col = 1
     If .Value = 1 Then
        .Col = 2
        vOperacion = .Text

        Call sbTraeFacturas
        Call sbTraeMontos(vOperacion)
     End If
   Next i
   

   
End With

Exit Sub
error:
  MsgBox Err.Description
  
End Sub

Private Sub sbTraeFacturas()
Dim vTotalFact As Double
Dim i As Integer
txtTotalFactura = 0
vTotalFact = 0

With vGridFacturas
 .MaxRows = 0
 strSQL = "Select DEPOSITO,NDOCUMENTO,FECHA_DOCUMENTO,DETALLE, MONTO " _
        & "from AFI_CD_DETALLE_LIQUIDACION where NOPERACION = " & vOperacion & " "
 rs.Open strSQL, glogon.Conection, adOpenStatic
 
 .MaxRows = 1
 
 Do While Not rs.EOF
    
    .Row = .MaxRows
    .Col = 1
    .Value = rs!DEPOSITO
    
    .Col = 2
    .Text = rs!nDocumento
    
    .Col = 3
    .Text = Format(rs!FECHA_DOCUMENTO, "dd/mm/yyyy")
    
    .Col = 4
    .Text = rs!Detalle
    
    .Col = 5
    .Text = Format(rs!Monto, "standard")
    vTotalFact = CDbl(vTotalFact) + Format(rs!Monto, "standard")
    
    rs.MoveNext
    .MaxRows = .MaxRows + 1
 
 Loop
 rs.Close
End With

End Sub


