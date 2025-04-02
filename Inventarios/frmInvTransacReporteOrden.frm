VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvTransacReporteOrden 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orden del Reporte"
   ClientHeight    =   4296
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7224
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvTransacReporteOrden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296
   ScaleWidth      =   7224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   492
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   4332
      _Version        =   1245187
      _ExtentX        =   7641
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Nombre del Producto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   492
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   4332
      _Version        =   1245187
      _ExtentX        =   7641
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Código del Producto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   492
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   4332
      _Version        =   1245187
      _ExtentX        =   7641
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Linea de Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   492
      Left            =   5160
      TabIndex        =   3
      Top             =   3480
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmInvTransacReporteOrden.frx":000C
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   732
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Orden del Detalle:"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvTransacReporteOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbReporte()
Dim vSQL As String, vSubTitulo As String, vOrden As String

vSQL = ""
vOrden = ""

Select Case InvTransacRep.Tipo
  Case "S"
     vSubTitulo = "SALIDAS"
  Case "E"
     vSubTitulo = "ENTRADAS"
  Case "T"
     vSubTitulo = "TRASLADOS"
  Case "R"
     vSubTitulo = "REQUISICION"
End Select

vSQL = "{PV_INVTRANSAC.BOLETA} = '" & InvTransacRep.Boleta & "' AND {PV_INVTRANSAC.TIPO} = '" & InvTransacRep.Tipo & "'"

Select Case True
  Case optX.Item(0).Value 'Descripcion
     vOrden = "+{PV_Productos.Descripcion}"
  Case optX.Item(1).Value 'Codigo
     vOrden = "+{pv_InvTraDet.cod_producto}"
  Case optX.Item(2).Value 'Linea
     vOrden = "+{pv_InvTraDet.Linea}"
End Select


Select Case InvTransacRep.Reporte
  Case "TrasladoC"
     Call sbInvReportes("TransaccionBoletaT", "BOLETA DE " & vSubTitulo, "", vSQL, vOrden)
  Case "TrasladoV"
     Call sbInvReportes("TransaccionBoletaTV", "BOLETA DE " & vSubTitulo & " PARA VENTA", "", vSQL, vOrden)
  Case Else
    Call sbInvReportes("TransaccionBoleta", "BOLETA DE " & vSubTitulo, "", vSQL, vOrden)
End Select


End Sub

Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

End Sub
