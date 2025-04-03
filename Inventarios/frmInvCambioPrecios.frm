VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvCambioPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Precios de Ventas"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   13332
   Icon            =   "frmInvCambioPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8304
   ScaleWidth      =   13332
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6852
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   13212
      _Version        =   1245187
      _ExtentX        =   23304
      _ExtentY        =   12086
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
      ItemCount       =   2
      Item(0).Caption =   "Listado de Excel"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "txtNoExisten"
      Item(0).Control(1)=   "txtCasos"
      Item(0).Control(2)=   "txtMonto"
      Item(0).Control(3)=   "vGridCarga"
      Item(0).Control(4)=   "txtArchivo"
      Item(0).Control(5)=   "Label2(2)"
      Item(0).Control(6)=   "Label2(1)"
      Item(0).Control(7)=   "Label2(3)"
      Item(0).Control(8)=   "Label1(2)"
      Item(0).Control(9)=   "cboPrecio"
      Item(0).Control(10)=   "Label2(7)"
      Item(0).Control(11)=   "chkExcel"
      Item(0).Control(12)=   "btnBuscar"
      Item(0).Control(13)=   "btnCargar"
      Item(0).Control(14)=   "btnInfo"
      Item(0).Control(15)=   "btnProcesar"
      Item(1).Caption =   "Factura de Proveedor"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "cmdCambiar"
      Item(1).Control(2)=   "lblProv"
      Item(1).Control(3)=   "lblFactura"
      Item(1).Control(4)=   "txtProvCod"
      Item(1).Control(5)=   "txtFactura"
      Item(1).Control(6)=   "txtProvDesc"
      Item(1).Control(7)=   "Label1(0)"
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Monto"
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox txtCasos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox txtNoExisten 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   6360
         Width           =   975
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4572
         Left            =   -69880
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   12972
         _Version        =   524288
         _ExtentX        =   22881
         _ExtentY        =   8064
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvCambioPrecios.frx":030A
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdCambiar 
         Height          =   612
         Left            =   -58720
         TabIndex        =   2
         Top             =   6120
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cambio de Precios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmInvCambioPrecios.frx":0ABF
      End
      Begin XtremeSuiteControls.FlatEdit txtProvCod 
         Height          =   312
         Left            =   -67840
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFactura 
         Height          =   312
         Left            =   -66400
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1245187
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProvDesc 
         Height          =   312
         Left            =   -66400
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1245187
         _ExtentX        =   10181
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   6120
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvCambioPrecios.frx":1482
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvCambioPrecios.frx":7CE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvCambioPrecios.frx":E546
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvCambioPrecios.frx":14DA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvCambioPrecios.frx":1B60A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridCarga 
         Height          =   4452
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   12852
         _Version        =   524288
         _ExtentX        =   22670
         _ExtentY        =   7853
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvCambioPrecios.frx":1B725
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   432
         Left            =   2040
         TabIndex        =   14
         Top             =   840
         Width           =   6852
         _Version        =   1245187
         _ExtentX        =   12086
         _ExtentY        =   762
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboPrecio 
         Height          =   312
         Left            =   2040
         TabIndex        =   19
         Top             =   480
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkExcel 
         Height          =   252
         Left            =   6840
         TabIndex        =   21
         Top             =   480
         Width           =   2052
         _Version        =   1245187
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Archivo Excel"
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   372
         Left            =   9120
         TabIndex        =   22
         Top             =   840
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmInvCambioPrecios.frx":1CA3F
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   372
         Left            =   9600
         TabIndex        =   23
         Top             =   840
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmInvCambioPrecios.frx":1D45D
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   372
         Left            =   10080
         TabIndex        =   24
         Top             =   840
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmInvCambioPrecios.frx":1DE20
      End
      Begin XtremeSuiteControls.PushButton btnProcesar 
         Height          =   612
         Left            =   9480
         TabIndex        =   25
         Top             =   6120
         Width           =   3132
         _Version        =   1245187
         _ExtentX        =   5524
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Procesar Cambios de Precio"
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
         Picture         =   "frmInvCambioPrecios.frx":1E5FF
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Precio:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   1200
         TabIndex        =   17
         Top             =   6360
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   3600
         TabIndex        =   16
         Top             =   6120
         Width           =   972
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No existen"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   4560
         TabIndex        =   15
         Top             =   6120
         Width           =   972
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Advertencia : Este proceso cambia el valor del precio regular de ventas del articulo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   612
         Index           =   0
         Left            =   -63040
         TabIndex        =   8
         Top             =   6120
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Label lblFactura 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "# Factura"
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
         Left            =   -67360
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label lblProv 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Left            =   -68920
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso de cambio de Precios del sistema"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   11412
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvCambioPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnBuscar_Click()
        txtArchivo.Text = ""
        
        With frmContenedor.CD
         If chkExcel.Value = vbChecked Then
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
                .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
                .ShowOpen
        
                If .FileName = "" Then
                    MsgBox "Archivo no válido...", vbExclamation
                    Exit Sub
                End If
        
                If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
                    'Ok
                Else
                    MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                    Exit Sub
                End If
        
                
         Else
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
                .Filter = "*.txt"
                .ShowOpen
                
                If .FileName = "" Then
                  MsgBox "Archivo no válido...", vbExclamation
                  Exit Sub
                End If
                
                If UCase(Right(.FileName, 3)) <> "TXT" Then
                  MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                  Exit Sub
                End If
         End If
        
         txtArchivo.Text = .FileName
        
        End With

End Sub

Private Sub btnCargar_Click()
    Call sbCargaArchivo
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CODIGO, PRECIO" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub


Private Sub btnProcesar_Click()
Dim strSQL As String, i   As Long

Dim pCodigo As String, pNombre As String, pMonto As Currency
Dim pTPrecio As String, pLinea As Long, pAnterior As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

pTPrecio = cboPrecio.ItemData(cboPrecio.ListIndex)

With vGridCarga

strSQL = ""
For i = 1 To .MaxRows
  .Row = i
  .col = 6
  If .Value = 1 Then
        .col = 1
        pCodigo = .Text
        .col = 5
        pMonto = CCur(.Text)
       strSQL = strSQL & Space(10) & " exec spInv_ListaPrecios_Procesa '" & pCodigo & "'," & pMonto & ",'" _
              & pTPrecio & "','" & glogon.Usuario & "'"
  End If
 
  If Len(strSQL) > 20000 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
  End If
 
Next i

  'Ultimo Lote
  If Len(strSQL) > 0 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
  End If


End With

Me.MousePointer = vbDefault
MsgBox "Actualizacion de Precios Finalizada...", vbInformation

Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cboPrecio_Click()
Call sbLimpia
End Sub

Private Sub cmdCambiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, i   As Integer
Dim curUtilidad As Currency, curCosto As Currency
Dim curPrecio As Currency, vPaso As Boolean
Dim curImpVentas As Currency, curPrecioNuevo As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

For i = 1 To vGrid.MaxRows
 vPaso = False
 vGrid.Row = i
 vGrid.col = 7
 If CCur(vGrid.Text) > 0 Then
   curPrecioNuevo = vGrid.Text
   vGrid.col = 3
   curUtilidad = vGrid.Text
   vGrid.col = 4
   curCosto = vGrid.Text
   vGrid.col = 1
    strSQL = "select (impuesto_ventas / 100) + 1 as Iv from pv_productos" _
           & " where cod_producto = '" & vGrid.Text & "'"
    Call OpenRecordSet(rs, strSQL)
        curImpVentas = rs!iv
    rs.Close
 
   Do While Not vPaso
    curPrecio = curCosto * ((curUtilidad / 100) + 1) * curImpVentas
   
    If curPrecio > curPrecioNuevo Then
      vPaso = True
      vGrid.col = 1
      strSQL = "update pv_productos set precio_regular = " & curCosto * ((curUtilidad / 100) + 1) _
             & ",porc_utilidad = " & curUtilidad _
             & " where cod_producto = '" & vGrid.Text & "'"
      Call ConectionExecute(strSQL)
    Else
      curUtilidad = curUtilidad + 0.001
    End If
    
    
   Loop
 
 End If
Next i

Me.MousePointer = vbDefault
MsgBox "Actualizacion Finalizada..."

Call sbCargaDatos

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 32

vGridCarga.MaxRows = 0

vGrid.MaxCols = 7
vGrid.MaxRows = 0

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vGrid.AppearanceStyle = fxGridStyle

tcMain.Item(0).Selected = True


strSQL = "select * from vInv_Tipos_Precios"
Call sbCbo_Llena_New(cboPrecio, strSQL, False, True)

Call sbCboAsignaDato(cboPrecio, "REGULAR", True, "TPR")

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbLimpia()
On Error GoTo vError
    
    
    tcMain.Item(0).Selected = True
    
    vGridCarga.MaxRows = "0"
    txtMonto.Text = "0"
    txtCasos.Text = "0"
    txtNoExisten.Text = "0"
    txtArchivo.Text = ""
    
   
vError:
End Sub



Private Sub sbCargaArchivo()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pCodigo As String, pNombre As String, pMonto As Currency
Dim pTPrecio As String, pLinea As Long, pAnterior As Currency

Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
Dim strMonto  As String
Dim strCedula As String
Dim strNombre As String
Dim i As Integer, vCampos As Boolean



On Error GoTo vError


vGridCarga.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


Me.MousePointer = vbHourglass


txtNoExisten.Text = 0
txtMonto.Text = 0
txtCasos.Text = 0

curMonto = 0
lCasos = 0 'Total

pTPrecio = cboPrecio.ItemData(cboPrecio.ListIndex)

Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CODIGO" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If vCampos Then
    vCampos = False
    For i = 0 To rs.Fields.Count
         
        If UCase(LCase(rs.Fields(i).Name)) = "PRECIO" Then
           vCampos = True
        End If
         
         If vCampos Then Exit For
    Next i


End If


If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son Codigo, Precio¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

'FIN: Validación del Archivo



'Sube, Revisa y Carga
With vGridCarga
    
    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Codigo) <> "" Then
        pCodigo = rs!Codigo
        pMonto = rs!Precio
        pLinea = pLinea + 1
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spInv_ListaPrecios_Sube '" & pCodigo & "'," & pMonto & ",'" _
                   & pTPrecio & "','" & glogon.Usuario & "'," & pLinea & "," & 1
        Else
            strSQL = strSQL & Space(10) & "exec spInv_ListaPrecios_Sube '" & pCodigo & "'," & pMonto & ",'" _
                   & pTPrecio & "','" & glogon.Usuario & "'," & pLinea & "," & 0
        End If
        
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If


'Revisa Lote y lo Carga
strSQL = "exec spInv_ListaPrecios_Consulta '" & pTPrecio & "','" & glogon.Usuario & "',1"
                   
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Exit Sub
End If

    Do While Not rs.EOF
            pCodigo = rs!Llave_01
            pNombre = rs!ref_01
            pMonto = rs!Monto_01
            pAnterior = rs!Monto_02
      
      
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = pCodigo
            .col = 2
            .Text = pNombre
            .col = 3
            .Value = IIf((rs!Detalle = "-1"), 0, 1)
            
            .col = 4
            .Text = Format(pAnterior, "Standard")
            
            .col = 5
            .Text = Format(pMonto, "Standard")
            
            
            If rs!Detalle = "-1" Then
               txtNoExisten.Text = CInt(txtNoExisten.Text) + 1
            Else
               .col = 6
               .Value = 1
            
            
                curMonto = curMonto + pMonto
            End If
            
            txtMonto.Text = Format(curMonto, "Standard")
            
            txtCasos = txtCasos + 1
            txtCasos.Refresh
      
      rs.MoveNext
    Loop
    rs.Close


End With 'vGrid



'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub




Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Len(txtFactura) > 0 Then
  Call sbCargaDatos
End If
End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "Select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = ""
  gBusquedas.Convertir = "N"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "Select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = ""
  gBusquedas.Convertir = "N"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub sbCargaDatos()
Dim strSQL As String

On Error GoTo vError

strSQL = "select P.cod_producto,P.descripcion,P.porc_utilidad,P.costo_regular" _
       & ",P.precio_regular * ((P.impuesto_ventas/100) + 1),D.precio,0" _
       & " from Cpr_Compras E inner join Cpr_Compras_detalle D on E.cod_factura = D.cod_factura" _
       & " and E.cod_proveedor = D.cod_proveedor" _
       & " inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " where E.cod_proveedor = " & txtProvCod _
       & " and E.cod_factura = '" & txtFactura & "'"
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL, True)
vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

