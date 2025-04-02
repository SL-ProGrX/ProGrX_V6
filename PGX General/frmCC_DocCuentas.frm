VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCC_DocCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indique la cuenta que desea utilizar para este tipo de documento"
   ClientHeight    =   4785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "frmCC_DocCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   3960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   16777215
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCC_DocCuentas.frx":000C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescribe 
      Height          =   312
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtReferencia 
      Height          =   312
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   6252
      _Version        =   1441793
      _ExtentX        =   11028
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   1632
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   6252
      _Version        =   1441793
      _ExtentX        =   11028
      _ExtentY        =   2879
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
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      Top             =   3960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BackColor       =   16777215
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCC_DocCuentas.frx":0733
      ImageAlignment  =   4
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique los datos para cerrar el documento como: Cuenta contable, referencias y notas para la nota."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5292
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmCC_DocCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxValidaCuenta(vCuenta As String) As Boolean
Dim vPaso As Boolean, rsX As New ADODB.Recordset, strSQL As String

 'Verifica que contablemente sea valida
 vPaso = fxgCntCuentaValida(vCuenta)
 
 'verifica que no sea utilizada en el auxiliar
 If vPaso Then
    strSQL = "exec spSIFValidaCuentas '" & Trim(vCuenta) & "'"
    Call OpenRecordSet(rsX, strSQL)
        vPaso = IIf((rsX!Existe = 0), True, False)
    rsX.Close
 End If
 
 fxValidaCuenta = vPaso
 
End Function


Private Sub btnTool_Click(Index As Integer)

Select Case Index
  Case 0 'Aceptar
        
        If Len(Trim(txtDetalle)) = 0 Then
          MsgBox "Ingrese el detalle de este documento ...", vbInformation
          Exit Sub
        End If
        
        txtDetalle.Text = fxCadenaDepura(txtDetalle.Text)
        txtReferencia.Text = Mid(fxCadenaDepura(txtReferencia.Text), 1, 30)
        
        If fxValidaCuenta(txtCuenta.Text) Then
         vAseDocDetalle = UCase(Mid(Trim(txtDetalle.Text), 1, 255))
         vAseDocDeposito = Trim(txtReferencia.Text)
         vAseDocCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
         vAseDocValido = True
         Unload Me
        Else
         MsgBox "La cuenta ingresada no es válida en el plan contable o Esta siendo utiliza por algun auxiliar, verifique...", vbCritical
         Form_Load
        End If
  
  Case 1 'Cancelar
     vAseDocValido = False
     Unload Me
End Select

End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vAseDocDetalle = ""
vAseDocDeposito = ""

txtDetalle.Text = ""
txtReferencia.Text = ""

vAseDocValido = False

txtCuenta.Text = vAseDocCuenta
txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 0)

txtDescribe.Text = fxgCntCuentaDesc(vAseDocCuenta)

End Sub


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 txtDetalle.SetFocus
 txtDescribe.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta.Text, 0))
 txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 1)
End If
If KeyCode = vbKeyF4 Then Call sbBuscaCuenta

End Sub

Private Sub sbBuscaCuenta()
Call sbgCntCuentaConsulta

If gBusquedas.Resultado <> "" Then
    txtCuenta.Text = fxgCntCuentaFormato(False, gBusquedas.Resultado, 1)
    txtDescribe.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta.Text, 0))
    txtCuenta.Text = fxgCntCuentaFormato(True, gBusquedas.Resultado, 1)
End If

End Sub


Function fxCadenaDepura(str As String) As String
Dim i As Integer, Resultado As String

str = Trim(str)
Resultado = ""

For i = 1 To Len(str)
 If Mid(str, i, 1) <> "'" Then Resultado = Resultado + Mid(str, i, 1)
Next i

fxCadenaDepura = Resultado

End Function

Private Sub txtDescribe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 txtReferencia.SetFocus
End If
End Sub

