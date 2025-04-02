VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCntX_CuentaForaneaInicializa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contabilidad"
   ClientHeight    =   4680
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   8835
   Icon            =   "frmCntX_CuentaForaneaInicializa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   8532
      _Version        =   1310722
      _ExtentX        =   15049
      _ExtentY        =   2138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnInicializa 
         Height          =   612
         Left            =   6960
         TabIndex        =   1
         Top             =   360
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Inicializa"
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
         Picture         =   "frmCntX_CuentaForaneaInicializa.frx":000C
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1560
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   312
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   5412
      _Version        =   1310722
      _ExtentX        =   9546
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
      Locked          =   -1  'True
      Appearance      =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAnio 
      Height          =   312
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   972
      _Version        =   1310722
      _ExtentX        =   1714
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtMes 
      Height          =   312
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   732
      _Version        =   1310722
      _ExtentX        =   1291
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtPeriodo 
      Height          =   312
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   5412
      _Version        =   1310722
      _ExtentX        =   9546
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
      Locked          =   -1  'True
      Appearance      =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtSaldoInicial 
      Height          =   312
      Left            =   1320
      TabIndex        =   9
      Top             =   2880
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
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
      Alignment       =   1
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtSI_DivisaFuncional 
      Height          =   312
      Left            =   1320
      TabIndex        =   12
      Top             =   2400
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
      _ExtentY        =   550
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin VB.Label lblDivisaCuenta 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   972
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDivisaFuncional 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   360
      TabIndex        =   14
      Top             =   2400
      Width           =   972
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Inicial en Divisa Funcional"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   2400
      Width           =   2532
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicializa Saldos en Divisa Extranjera"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   240
      Width           =   5892
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Inicial en Divisa Extranjera"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2880
      Width           =   2532
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   972
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   972
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10572
   End
End
Attribute VB_Name = "frmCntX_CuentaForaneaInicializa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDivisaFuncional As String


Private Sub btnInicializa_Click()

On Error GoTo vError

'Verifica
If txtCuenta.Text = "" Or Not IsNumeric(txtSaldoInicial.Text) Then
  MsgBox "Datos Incorrectos Verifique!", vbInformation
  Exit Sub
End If

Dim i As Integer

i = MsgBox("Esta seguro de realizar el Ajuste al Saldo Inicial en divisa extranjera?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


Me.MousePointer = vbHourglass


'Realiza el Proceso
strSQL = "exec spCntX_Cuenta_Foranea_Inicializa " & gCntX_Parametros.CodigoConta _
        & ", '', '" & txtCuenta.Text & "', " & CCur(txtSaldoInicial.Text) _
        & ", " & txtAnio.Text & ", " & txtMes.Text & ", '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

'Bitacora
Call Bitacora("Inicializa", "Saldo Divisa Extranjera, Cuenta: " & txtCuenta.Text _
        & ", Saldo Inicial: " & txtSaldoInicial.Text)

Me.MousePointer = vbDefault

MsgBox "Cambio Realizado Saltisfactoriamente!", vbInformation


Exit Sub

vError:
Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 


End Sub

Private Sub sbLimpia()

txtCuenta.Text = ""
txtCuentaDesc.Text = ""

txtSaldoInicial.Text = 0
txtSI_DivisaFuncional.Text = 0

lblDivisaCuenta.Caption = ""


End Sub

Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture


txtAnio.Text = gCntX_Parametros.PeriodoAnio
txtMes.Text = gCntX_Parametros.PeriodoMes

txtPeriodo.Text = fxCntX_PeriodoDesc(txtAnio, txtMes)


strSQL = "select cod_divisa from CntX_Divisas " _
       & " Where DIVISA_LOCAL = 1 and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL)
 vDivisaFuncional = RTrim(rs!COD_DIVISA)
rs.Close

lblDivisaFuncional.Caption = vDivisaFuncional
Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
     gBusquedas.Resultado = ""
     gBusquedas.Resultado2 = ""
     gBusquedas.Col1Name = "Cuenta"
     gBusquedas.Col2Name = "Descripción"
     gBusquedas.Col3Name = "Divisa"
     gBusquedas.Consulta = "select cod_cuenta_mask, descripcion, cod_divisa from CntX_Cuentas"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and cod_divisa <> '" & vDivisaFuncional & "' and Acepta_Movimientos = 1"
     gBusquedas.Columna = "cod_cuenta_Mask"
     gBusquedas.Orden = "cod_cuenta_Mask"
     frmBusquedas.Show vbModal
     
     If gBusquedas.Resultado <> "" Then
        txtCuenta.Text = gBusquedas.Resultado
        txtCuentaDesc.Text = gBusquedas.Resultado2
        lblDivisaCuenta.Caption = gBusquedas.Resultado3
        
        strSQL = "select Saldo_Inicial, DF_Saldo_Inicial" _
                & " From vCntX_Mov_Cuentas_General" _
                & "    Where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                & "      and cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "'" _
                & "      and anio = " & txtAnio & " and mes = " & txtMes.Text
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
            txtSI_DivisaFuncional.Text = Format(rs!Saldo_Inicial, "Standard")
            txtSaldoInicial.Text = Format(rs!DF_Saldo_Inicial, "Standard")
        End If
        rs.Close
     Else
        Call sbLimpia
     End If
End If
End Sub
