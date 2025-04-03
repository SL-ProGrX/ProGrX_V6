VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Prendas_Old 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prendas de la Operación"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   3201
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
      FlatScrollBar   =   -1  'True
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9960
      Top             =   600
   End
   Begin XtremeSuiteControls.GroupBox fra 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   6588
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtModelo 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   2520
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtSerie 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2880
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   915
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   7815
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   1614
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
      Begin XtremeSuiteControls.FlatEdit txtCobertura 
         Height          =   330
         Left            =   7320
         TabIndex        =   15
         Top             =   2880
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMarca 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   2160
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAvaluo 
         Height          =   330
         Left            =   7320
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   582
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCoberturaPorc 
         Height          =   315
         Left            =   7320
         TabIndex        =   14
         Top             =   2520
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "% Cobertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   17
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cobertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serie/Año"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Avalúo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   330
      Left            =   6240
      TabIndex        =   12
      Top             =   720
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCoberturaTotal 
      Height          =   450
      Left            =   7320
      TabIndex        =   18
      Top             =   7800
      Width           =   2895
      _Version        =   1572864
      _ExtentX        =   5106
      _ExtentY        =   794
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   1320
      TabIndex        =   21
      Top             =   720
      Width           =   4815
      _Version        =   1572864
      _ExtentX        =   8493
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   23
      Top             =   1080
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Registro de Prendas:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   1095
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10335
      _Version        =   1572864
      _ExtentX        =   18230
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Prendas para la Operación: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cobertura Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   19
      Top             =   7800
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scRegistrado 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   5520
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Prendas Registradas:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
End
Attribute VB_Name = "frmCR_Prendas_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mPrendaId As Long

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vEdita As Integer, mFecha As Date



Private Sub Form_Load()

scMain.Caption = "Prendas para la Operación: " & Operacion.Operacion

With lsw.ColumnHeaders
    .Clear
    .Add , , "Tipo", 1000
    .Add , , "Categoria", 2500
    .Add , , "Avaluo", 1800, vbRightJustify
    .Add , , "%", 1800, vbRightJustify
    .Add , , "Cobertura", 1800, vbRightJustify
    .Add , , "Descripción", 2500
    .Add , , "Modelo", 2500
    .Add , , "Serie", 2500
    .Add , , "Marca", 2500
End With

Call sbToolBarIconos(tlbPrincipal, False)

With tlbPrincipal
    .Buttons(1).Enabled = True
    .Buttons(2).Enabled = False
    .Buttons(3).Enabled = False
    .Buttons(4).Enabled = False
    .Buttons(5).Enabled = False
End With

fra.Enabled = False

End Sub

Private Sub sbLimpia()
 txtDescripcion.Text = ""
 txtModelo.Text = ""
 txtSerie.Text = ""
 txtMarca.Text = ""
 
 txtAvaluo.Text = "0"
 txtCoberturaPorc.Text = "0"
 txtCobertura.Text = "0"
 
 mPrendaId = 0
 
End Sub

Private Sub sbPrendas_Load()

Dim curTotal As Currency

curTotal = 0

strSQL = "exec spCrd_Operacion_Prenda_Consulta " & Operacion.Operacion

Call OpenRecordSet(rs, strSQL)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Tipo_Prenda)
      
      itmX.SubItems(1) = rs!PrendaDesc
      
      
      itmX.SubItems(2) = Format(rs!Avaluo, "Standard")
      itmX.SubItems(3) = Format(rs!Porc_Cobertura, "Standard")
      itmX.SubItems(4) = Format(rs!Cobertura, "Standard")
      
      itmX.SubItems(5) = rs!Descripcion
      itmX.SubItems(6) = rs!Modelo
      itmX.SubItems(7) = rs!Serie
      itmX.SubItems(8) = rs!Marca
      
      itmX.Tag = rs!Prenda_Id
      
      curTotal = curTotal + rs!Cobertura
 rs.MoveNext
Loop
rs.Close

txtCoberturaTotal.Text = Format(curTotal, "Standard")

End Sub




Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

 mPrendaId = Item.Tag
 
 strSQL = "exec spCrd_Operacion_Prenda_Consulta " & Operacion.Operacion & ", " & mPrendaId
 
 Call OpenRecordSet(rs, strSQL)
 
 
 txtDescripcion.Text = rs!Descripcion & ""
 txtModelo.Text = rs!Modelo & ""
 txtSerie.Text = rs!Serie & ""
 txtMarca.Text = rs!Marca & ""
 
 txtAvaluo.Text = Format(rs!Avaluo, "Standard")
 txtCoberturaPorc.Text = Format(rs!Porc_Cobertura, "Standard")
 txtCobertura.Text = Format(rs!Cobertura, "Standard")
 
 Call sbCboAsignaDato(cboTipo, rs!PrendaDesc, True, Trim(rs!Tipo_Prenda))
 

With tlbPrincipal
   .Buttons(1).Enabled = False
   .Buttons(2).Enabled = True
   .Buttons(3).Enabled = True
   .Buttons(4).Enabled = False
   .Buttons(5).Enabled = False
End With


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbInicializa()

On Error GoTo vError

strSQL = "select rtrim(tipo_prenda) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from crd_prendas_tipos where Activa = 1 order by descripcion "

Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
 

Call sbPrendas_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub


Private Sub txtAvaluo_GotFocus()
On Error GoTo vError
    txtAvaluo.Text = CCur(txtAvaluo.Text)
vError:

End Sub

Private Sub txtAvaluo_LostFocus()
On Error GoTo vError
    txtAvaluo.Text = Format(CCur(txtAvaluo.Text), "Standard")
vError:
End Sub


Private Sub txtAvaluo_KeyPress(KeyAscii As Integer)
On Error GoTo vError
  If KeyAscii = vbKeyReturn Then txtCoberturaPorc.SetFocus
vError:
End Sub

Private Function fxVerificaDatos() As Boolean
Dim vMensaje As String

fxVerificaDatos = True
vMensaje = ""

'Revision de Inyección

txtDescripcion.Text = fxSysCleanTxtInject(txtDescripcion.Text)
txtModelo.Text = fxSysCleanTxtInject(txtModelo.Text)
txtSerie.Text = fxSysCleanTxtInject(txtSerie.Text)
txtMarca.Text = fxSysCleanTxtInject(txtMarca.Text)



If Len(Trim(txtDescripcion)) < 10 Then vMensaje = vMensaje & vbCrLf & "- La descripción no es válida"
If Not IsNumeric(txtAvaluo.Text) Then
   vMensaje = vMensaje & vbCrLf & "- El dato del Avalúo es erroneo!"
End If


If Len(vMensaje) > 0 Then
  fxVerificaDatos = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuarda()

On Error GoTo vError
        
 strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & Operacion.Operacion & ", '" & Operacion.Codigo & "', '" _
        & cboTipo.ItemData(cboTipo.ListIndex) & "', " & CCur(txtAvaluo.Text) & ", '" & txtDescripcion.Text & "', '" & txtModelo.Text & "', '" & txtSerie.Text _
        & "', '" & txtMarca.Text & "', '" & glogon.Usuario & "', 'A'"

 Call ConectionExecute(strSQL)
 
 MsgBox "Prenda registrada satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer

Select Case Button.Key
  Case "insertar", "nuevo"
   vEdita = 0
   Call sbLimpia
    
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "editar", "modificar"
   vEdita = 1
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "borrar"
   
   If mPrendaId > 0 Then
    iRespuesta = MsgBox("Esta seguro que desea eliminar esta prenda?", vbYesNo)
    
    If iRespuesta = vbYes Then
       strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & Operacion.Operacion & ", '" & Operacion.Codigo & "', '" _
        & cboTipo.ItemData(cboTipo.ListIndex) & "', " & CCur(txtAvaluo.Text) & ",'" & txtDescripcion.Text & "', '" & txtModelo.Text & "', '" & txtSerie.Text _
        & "', '" & txtMarca.Text & "', '" & glogon.Usuario & "', 'E'"


      Call ConectionExecute(strSQL)
      Call sbPrendas_Load
      Call sbLimpia
    Else
      Call sbLimpia
    End If
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
     End With
    
   End If
  
  Case "salvar", "guardar"
    If fxVerificaDatos Then
      Call sbGuarda
      
      Call sbPrendas_Load
      
      With tlbPrincipal
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
      End With
      
      Call sbLimpia
    
    Else
      MsgBox "Información Ingresada es Incorrecta por favor verifique...", vbInformation
    End If
  
  Case "deshacer"
    Call sbLimpia
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
    End With
  
  Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
        
  Case "salir", "cerrar"
    Unload Me

End Select

End Sub




