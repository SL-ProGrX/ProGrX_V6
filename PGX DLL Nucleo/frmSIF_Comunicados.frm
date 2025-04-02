VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmSIF_Comunicados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comunicados"
   ClientHeight    =   6132
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8208
   Icon            =   "frmSIF_Comunicados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6132
   ScaleWidth      =   8208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkCursiva 
      Height          =   252
      Left            =   1440
      TabIndex        =   22
      Top             =   4320
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cursiva"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin VB.TextBox txtMuestra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Text            =   "-->> Texto de Demostración de Comunicados <<--"
      Top             =   4800
      Width           =   6492
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   612
      Left            =   6480
      TabIndex        =   12
      Top             =   5400
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmSIF_Comunicados.frx":000C
   End
   Begin XtremeSuiteControls.PushButton cmdNuevo 
      Height          =   612
      Left            =   5040
      TabIndex        =   13
      Top             =   5400
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmSIF_Comunicados.frx":0703
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2400
      TabIndex        =   14
      Top             =   3240
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   2400
      TabIndex        =   15
      Top             =   3600
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtNota 
      Height          =   1512
      Left            =   1440
      TabIndex        =   16
      Top             =   1560
      Width           =   6492
      _Version        =   1245187
      _ExtentX        =   11451
      _ExtentY        =   2667
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1440
      TabIndex        =   17
      Top             =   1080
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   5640
      TabIndex        =   18
      Top             =   3240
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   312
      Left            =   5640
      TabIndex        =   19
      Top             =   3600
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboFuente 
      Height          =   312
      Left            =   1440
      TabIndex        =   20
      Top             =   3960
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
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
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboColor 
      Height          =   312
      Left            =   5640
      TabIndex        =   21
      Top             =   3960
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
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
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkNegrita 
      Height          =   252
      Left            =   2760
      TabIndex        =   23
      Top             =   4320
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Negrita"
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
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   9
      Left            =   4440
      TabIndex        =   24
      Top             =   3960
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   4440
      TabIndex        =   11
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Muestra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   3
      Left            =   4440
      TabIndex        =   4
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vigencia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   0
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comunicados del área de servicios"
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
      Height          =   492
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11652
   End
End
Attribute VB_Name = "frmSIF_Comunicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean

Private Sub sbConsulta(vCom As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Call sbLimpia

strSQL = "select * from sif_comunicados where cod_comunicado = " & vCom
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtCodigo = rs!cod_comunicado
    txtCodigo.Tag = 0
    txtCodigo.ToolTipText = rs!cod_comunicado
    
    txtUsuario.Text = rs!Usuario
    txtFecha.Text = rs!fecha

    txtNota.Text = rs!nota
    dtpInicio.Value = rs!Inicio
    dtpCorte.Value = rs!Corte

    Select Case rs!fColor
      Case "A"
        cboColor.Text = "Azul"
      Case "N"
        cboColor.Text = "Negro"
      Case "V"
        cboColor.Text = "Verde"
      Case "R"
        cboColor.Text = "Rojo"
    End Select
    
    chkCursiva.Value = rs!fCursiva
    chkNegrita.Value = rs!fNegrita
    
    cboFuente.Text = rs!fFuente

End If
rs.Close

End Sub



Private Sub cboColor_Click()
Call sbMuestra
End Sub

Private Sub cboFuente_Click()
Call sbMuestra
End Sub

Private Sub chkCursiva_Click()
Call sbMuestra
End Sub

Private Sub chkNegrita_Click()
Call sbMuestra
End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xComunicado As Long

On Error GoTo vError

If Len(txtNota) = 0 Then Exit Sub

If txtCodigo.Tag = "1" Then
   strSQL = "insert sif_comunicados(usuario,fecha,inicio,corte,nota,fFuente,fColor,fCursiva,fnegrita)" _
          & " values('" & glogon.Usuario & "',dbo.MyGetdate(),'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
          & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" & txtNota.Text & "','" & cboFuente.Text _
          & "','" & Mid(cboColor.Text, 1, 1) & "'," & chkCursiva.Value & "," & chkNegrita.Value & ")"
   Call ConectionExecute(strSQL)
   
   strSQL = "select max(cod_comunicado) as Ultimo from sif_comunicados"
   Call OpenRecordSet(rs, strSQL)
    xComunicado = rs!Ultimo
   rs.Close
   
Else
  strSQL = "update sif_comunicados set usuario = '" & glogon.Usuario & "',fecha = dbo.MyGetdate(), inicio = '" _
          & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00',corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
          & " 23:59:59',nota = '" & txtNota.Text & "',fFuente = '" & cboFuente.Text & "', fColor = '" _
          & Mid(cboColor.Text, 1, 1) & "',fCursiva = " & chkCursiva.Value & ", fNegrita = " & chkNegrita.Value _
          & " Where cod_comunicado = " & txtCodigo.ToolTipText
  Call ConectionExecute(strSQL)
  
  xComunicado = txtCodigo.ToolTipText
  
End If

MsgBox "Comunicado # " & xComunicado & " Registrado Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdNuevo_Click()
 Call sbLimpia
 txtNota.SetFocus
End Sub

Private Sub sbLimpia()
 txtCodigo = "0"
 txtCodigo.Tag = 1
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value
 
 txtNota.Text = ""
 txtUsuario.Text = ""
 txtFecha.Text = ""
 
 cboColor.Text = "Azul"
 cboFuente.Text = "Courier New"
 
 chkCursiva.Value = vbUnchecked
 chkNegrita.Value = vbUnchecked
 
End Sub


Private Sub sbMuestra()

On Error GoTo vError

txtMuestra.FontBold = IIf((chkNegrita.Value = vbChecked), True, False)
txtMuestra.FontItalic = IIf((chkCursiva.Value = vbChecked), True, False)
txtMuestra.FontName = cboFuente.Text

Select Case Mid(cboColor.Text, 1, 1)
  Case "N"
    txtMuestra.ForeColor = vbBlack
  Case "V"
    txtMuestra.ForeColor = vbGreen
  Case "R"
    txtMuestra.ForeColor = vbRed
  Case "A"
    txtMuestra.ForeColor = vbBlue
  Case Else
    txtMuestra.ForeColor = vbBlack
End Select

vError:


End Sub

Private Sub Form_Load()
 vModulo = 10
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 cboColor.Clear
 cboColor.AddItem "Azul"
 cboColor.AddItem "Negro"
 cboColor.AddItem "Rojo"
 cboColor.AddItem "Verde"
 
 cboFuente.Clear
 cboFuente.AddItem "Arial"
 cboFuente.AddItem "Arial Narrow"
 cboFuente.AddItem "Courier New"
 cboFuente.AddItem "Tahoma"
 cboFuente.AddItem "Times New Roman"
 
 
 Call sbLimpia
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta(txtCodigo)

vError:
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_comunicado from sif_comunicados"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_comunicado > " & txtCodigo & " order by cod_comunicado asc"
    Else
       strSQL = strSQL & " where cod_comunicado < " & txtCodigo & " order by cod_comunicado desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_comunicado
      Call sbConsulta(rs!cod_comunicado)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


