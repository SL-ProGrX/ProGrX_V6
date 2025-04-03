VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCprOrdenesProceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orden de Compra [Procesos]"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTipo 
      Alignment       =   2  'Center
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   4812
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4095
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   10695
      _Version        =   1441793
      _ExtentX        =   18865
      _ExtentY        =   7223
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
      Item(0).Caption =   "Cotización"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "txtC_ProvDesc"
      Item(0).Control(1)=   "txtC_ProvId"
      Item(0).Control(2)=   "Label2(2)"
      Item(0).Control(3)=   "btnCotizacion(0)"
      Item(0).Control(4)=   "btnCotizacion(1)"
      Item(0).Control(5)=   "btnCotizacion(2)"
      Item(0).Control(6)=   "lswCotizacion"
      Item(0).Control(7)=   "btnCotizacion(3)"
      Item(0).Control(8)=   "btnCotizacion(4)"
      Item(0).Control(9)=   "ShortcutCaption1"
      Item(1).Caption =   "Orden de Compra"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "txtPin"
      Item(1).Control(1)=   "chkPinIngreso"
      Item(1).Control(2)=   "txtProvDescripcion"
      Item(1).Control(3)=   "txtProvCodigo"
      Item(1).Control(4)=   "imgPinGenera"
      Item(1).Control(5)=   "Label2(0)"
      Item(1).Control(6)=   "Label2(1)"
      Item(1).Control(7)=   "btnPIN"
      Item(1).Control(8)=   "bntOrdenCompra"
      Item(2).Caption =   "Estado"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "btnCerrar"
      Item(2).Control(1)=   "Label2(4)"
      Item(2).Control(2)=   "btnResolucion(0)"
      Item(2).Control(3)=   "btnResolucion(1)"
      Item(2).Control(4)=   "Label2(5)"
      Begin XtremeSuiteControls.ListView lswCotizacion 
         Height          =   2055
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   3625
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkPinIngreso 
         Height          =   372
         Left            =   -64960
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   3252
         _Version        =   1441793
         _ExtentX        =   5736
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Utilizar Pin para Ingreso de Mercadería"
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
         Appearance      =   16
         Alignment       =   1
      End
      Begin VB.TextBox txtProvCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -68200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtProvDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -67000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   5292
      End
      Begin VB.TextBox txtPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -63640
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin XtremeSuiteControls.PushButton btnCotizacion 
         Height          =   375
         Index           =   0
         Left            =   8760
         TabIndex        =   8
         Top             =   480
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483633
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
         Appearance      =   17
         Picture         =   "frmCrpOrdenesProceso.frx":0000
      End
      Begin VB.TextBox txtC_ProvId 
         Alignment       =   2  'Center
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtC_ProvDesc 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   6012
      End
      Begin XtremeSuiteControls.PushButton btnCotizacion 
         Height          =   375
         Index           =   1
         Left            =   6600
         TabIndex        =   9
         Top             =   840
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnCotizacion 
         Height          =   375
         Index           =   2
         Left            =   7680
         TabIndex        =   10
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnCotizacion 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cotización"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnCotizacion 
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   13
         Top             =   840
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Adjudicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnPIN 
         Height          =   375
         Left            =   -63640
         TabIndex        =   19
         Top             =   1890
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reemplazar PIN"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton bntOrdenCompra 
         Height          =   615
         Left            =   -64360
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Emitir Orden de Compra"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCrpOrdenesProceso.frx":0700
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   612
         Left            =   -63760
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cerrar la Orden"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCrpOrdenesProceso.frx":0E19
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   612
         Index           =   0
         Left            =   -63760
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Autorizar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCrpOrdenesProceso.frx":1805
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   612
         Index           =   1
         Left            =   -63760
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Rechazar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCrpOrdenesProceso.frx":1FE3
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   1440
         TabIndex        =   28
         Top             =   1560
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado de Proveedores consultados:"
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
      Begin VB.Label Label2 
         Caption         =   $"frmCrpOrdenesProceso.frx":2978
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Index           =   5
         Left            =   -69760
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   5532
      End
      Begin VB.Label Label2 
         Caption         =   $"frmCrpOrdenesProceso.frx":2A4D
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Index           =   4
         Left            =   -69640
         TabIndex        =   24
         Top             =   2640
         Visible         =   0   'False
         Width           =   5532
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   -69400
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "Pin"
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
         Left            =   -64000
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Image imgPinGenera 
         Height          =   252
         Left            =   -61600
         Picture         =   "frmCrpOrdenesProceso.frx":2B30
         Stretch         =   -1  'True
         ToolTipText     =   "Genera Pin"
         Top             =   1560
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
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
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1212
      End
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   6612
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCprOrdenesProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mEstado As String, mProceso As String, mOrden As String

Private Sub bntOrdenCompra_Click()
   
   MsgBox "Se a enviado correo al Proveedor con la Orden de Compra!", vbInformation
   
   Call sbReporte_Orden_Compra

End Sub

Private Sub btnCerrar_Click()
Dim strSQL As String

On Error GoTo vError

If mEstado = "A" And (mProceso = "A" Or mProceso = "D" Or mProceso = "Y") Then
    strSQL = "update CPR_Ordenes set Proceso = 'Y' " _
           & " where cod_Orden = '" & txtCodigo.Text & "' and Estado = 'A' and Proceso in('A','D','X')"
    Call ConectionExecute(strSQL)
    
    Call sbInicializa
    
    MsgBox "Orden Cerrada satisfactoriamente!", vbInformation

Else
    MsgBox "La Orden no puede ser cerrada porque el proceso o estado actual no lo permiten!", vbExclamation
     
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxCxP_Proveedor_Estado(pProveedor As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Dim vResult As Boolean

vResult = True

strSQL = "select Estado from CxP_Proveedores where cod_proveedor = " & pProveedor
Call OpenRecordSet(rs, strSQL)

If rs!Estado <> "A" Then
    vResult = False
End If


fxCxP_Proveedor_Estado = vResult

End Function

Private Sub btnCotizacion_Click(Index As Integer)
Dim strSQL As String

If Not IsNumeric(txtC_ProvId.Text) And Index <> 0 Then
   MsgBox "No. de Proveedor no es válido!", vbExclamation
   Exit Sub
End If

If Not fxCxP_Proveedor_Estado(txtC_ProvId) Then
    MsgBox "El Proveedor no se encuentra activo, verifique el estado!", vbExclamation
    Exit Sub
End If


Select Case Index
  Case 0 'Consulta
        Call txtC_ProvId_KeyDown(vbKeyF4, 0)
        Exit Sub
  Case 1 'Agregar
        strSQL = "exec spCpr_Orden_Proceso '" & txtCodigo.Text & "','" & glogon.Usuario & "','+'," & txtC_ProvId.Text & ",''"
        Call ConectionExecute(strSQL)
        
  Case 2 'Borrar
        strSQL = "exec spCpr_Orden_Proceso '" & txtCodigo.Text & "','" & glogon.Usuario & "','-'," & txtC_ProvId.Text & ",''"
        Call ConectionExecute(strSQL)
 
  
  Case 3 'Cotizar
        strSQL = "exec spCpr_Orden_Proceso '" & txtCodigo.Text & "','" & glogon.Usuario & "','C'," & txtC_ProvId.Text & ",''"
        Call ConectionExecute(strSQL)
        If Not glogon.error Then
        
            MsgBox "Se enviado Solicitud de Cotización al Proveedor: " & txtC_ProvDesc, vbInformation
        
            Call sbInicializa
            Call sbReporte_Cotizacion
        End If
        
  Case 4 'Adjudicar
        strSQL = "exec spCpr_Orden_Proceso '" & txtCodigo.Text & "','" & glogon.Usuario & "','A'," & txtC_ProvId.Text & ",''"
        Call ConectionExecute(strSQL)

        Call sbInicializa

End Select

Call sbLista_Proveedores

End Sub

Private Sub btnPIN_Click()
 Call sbPin
End Sub

Private Sub sbLista_Proveedores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

strSQL = "select Prov.COD_PROVEEDOR, Prov.DESCRIPCION, Op.*" _
       & " from CPR_ORDENES O inner join CPR_ORDENES_PROCESO Op on O.COD_ORDEN = Op.COD_ORDEN" _
       & "                    inner join CXP_PROVEEDORES Prov on Op.cod_proveedor = Prov.COD_PROVEEDOR" _
       & " WHERE O.COD_ORDEN = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)

lswCotizacion.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswCotizacion.ListItems.Add(, , rs!cod_proveedor)
     itmX.SubItems(1) = rs!DESCRIPCION
     itmX.SubItems(2) = rs!Cotiza_Fecha & ""
     itmX.SubItems(3) = rs!Adjudica_Fecha & ""
     itmX.SubItems(4) = rs!REGISTRO_FECHA & ""
     itmX.SubItems(5) = rs!REGISTRO_USUARIO & ""
     itmX.SubItems(6) = rs!Cotiza_Usuario & ""
     itmX.SubItems(7) = rs!Adjudica_Usuario & ""
 
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True

lswCotizacion.ListItems.Clear

      
strSQL = "select O.Estado, O.Proceso, P.cod_proveedor,P.descripcion,O.pin_autorizacion, C.descripcion as 'Causa_Desc'" _
       & " from cpr_ordenes O  inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & "  left join cxp_proveedores P On O.cod_proveedor = P.cod_proveedor" _
       & " where cod_orden = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  
  txtTipo.Text = rs!Causa_Desc
  
  txtProvCodigo = rs!cod_proveedor & ""
  txtProvDescripcion = rs!DESCRIPCION & ""
  
  chkPinIngreso.Value = rs!pin_autorizacion
  
  mEstado = rs!Estado
  mProceso = rs!Proceso

  Select Case rs!Estado
    Case "S"
      txtEstado = "Solicitada"
    Case "A"
      txtEstado = "Autorizada"
    Case "R"
      txtEstado = "Rechazada"
    Case Else
      txtEstado = "No Identificada"
  End Select

  Select Case rs!Proceso
    Case "P"
      txtEstado = txtEstado & " ¦ Pendiente"
    Case "C"
      txtEstado = txtEstado & " ¦ Cotizada"
    Case "A"
      txtEstado = txtEstado & " ¦ Adjudicada"
    Case "D"
      txtEstado = txtEstado & " ¦ Despacho Total"
    Case "X"
      txtEstado = txtEstado & " ¦ Despacho Parcial"
    Case "Y"
      txtEstado = txtEstado & " ¦ Cerrada"
    Case Else
      txtEstado = txtEstado & " ¦ No Identificada"
  End Select
  
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnResolucion_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

If mEstado <> "S" Then
  MsgBox "El Estado de la Orden no se encuentra pendiente, ya ha sido resuelta en el pasado!", vbExclamation
  Exit Sub
End If


'Verifica que el usuario tenga autorizacion para anular

strSQL = "select count(*) as 'Existe'" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " where O.autoriza_fecha is null and O.estado = 'S'" _
       & "   and O.genera_user in(" _
       & "select usuario_asignado from cpr_orden_autousers where usuario = '" _
       & glogon.Usuario & "')"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  MsgBox "La orden no esta pendiente o el usuario actual no está autorizado para esta gestión!", vbExclamation
  Exit Sub
End If
rs.Close


If Index = 0 Then
    strSQL = "update cpr_ordenes set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
           & glogon.Usuario & "',estado = 'A' where cod_orden = '" & txtCodigo.Text & "'"

  Call Bitacora("Aplica", "Orden de Compra Id: " & txtCodigo.Text & ", Autorización!")

Else
  strSQL = "update cpr_ordenes set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
         & glogon.Usuario & "',estado = 'R' where cod_orden = '" & txtCodigo.Text & "'"
  
  Call Bitacora("Aplica", "Orden de Compra Id: " & txtCodigo.Text & ", Rechazo!")

End If

Call ConectionExecute(strSQL)

Call sbInicializa
MsgBox "Resolución aplicada satisfactoriamente!", vbInformation

End Sub

Private Sub Form_Load()
vModulo = 35

txtCodigo = GLOBALES.gTag

With lswCotizacion.ColumnHeaders
  .Clear
  .Add , , "Prov.Id", 1200, vbCenter
  .Add , , "Descripción", 3600
  .Add , , "Cotización", 2000, vbCenter
  .Add , , "Adjudicado?", 2000, vbCenter
  .Add , , "Registro", 2000, vbCenter
  .Add , , "Us.Registro", 1200
  .Add , , "Us.Cotiza", 1200
  .Add , , "Us.Adjudica", 1200
End With

Call sbInicializa
Call sbLista_Proveedores

End Sub

Private Sub imgPinGenera_Click()
Dim i As Integer

txtPin = Mid(glogon.Usuario, 1, 1) & Format(Day(Date), "00") & Right(glogon.Usuario, 1)
txtPin = Format(Hour(Time), "00") & txtPin & Format(Second(Time), "00")

Randomize
i = Int((30 * Rnd) + 1)

txtPin = Chr(64 + i) & txtPin

Randomize
i = Int((30 * Rnd) + 1)

txtPin = txtPin & Chr(64 + i)


End Sub

Private Sub sbPin()
Dim strSQL As String

On Error GoTo vError

Select Case mProceso
  Case "D", "X", "Y"
          MsgBox "La orden se encuentra en proceso de despacho, no puede ser modificado el Pin", vbInformation
  
  Case Else
        strSQL = "update cpr_ordenes set pin_entrada = '" & txtPin & "',pin_autorizacion = " & chkPinIngreso.Value _
               & " where cod_orden = '" & txtCodigo & "' and proceso not in('D','X')"
        Call ConectionExecute(strSQL)
        
        MsgBox "Pin reemplazado Satisfactoriamente...", vbInformation
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReporte_Orden_Compra()
Dim vSQL As String

vSQL = "{cpr_ordenes.COD_ORDEN} = '" & txtCodigo & "'"

'Solicitar primer Proveedor
If txtProvCodigo <> "" Then
   Call sbInvReportes("OrdenesCompra", "ORDEN DE COMPRA", txtProvDescripcion.Text, vSQL)
End If

End Sub


Private Sub sbReporte_Cotizacion()
Dim vSQL As String, strSQL As String, rs As New ADODB.Recordset

vSQL = "{cpr_ordenes.COD_ORDEN} = '" & txtCodigo & "'"

Call sbInvReportes("OrdenesCotizacion", "SOLICITUD DE COTIZACION", UCase(txtC_ProvDesc.Text), vSQL)

End Sub

Private Sub lswCotizacion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

txtC_ProvId.Text = Item.Text
txtC_ProvDesc.Text = Item.SubItems(1)

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "pin"
    Call sbPin
  Case "reporte"
    Call sbReporte_Cotizacion
End Select
End Sub


Private Sub txtC_ProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtC_ProvId.Text = gBusquedas.Resultado
  txtC_ProvDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtC_ProvId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtC_ProvId.Text = gBusquedas.Resultado
  txtC_ProvDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvDescripcion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtProvCodigo = gBusquedas.Resultado
  txtProvDescripcion = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkPinIngreso.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtProvCodigo = gBusquedas.Resultado
  txtProvDescripcion = gBusquedas.Resultado2
End If

End Sub
