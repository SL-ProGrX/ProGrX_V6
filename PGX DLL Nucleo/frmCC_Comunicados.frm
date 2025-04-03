VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSIF_Comunicados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicados"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCC_Comunicados.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   8160
   Begin VB.TextBox txtMuestra 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1560
      TabIndex        =   22
      Text            =   "-->> Texto de Demostración de Comunicados <<--"
      Top             =   4320
      Width           =   6375
   End
   Begin VB.ComboBox cboFuente 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CheckBox chkNegrita 
      Appearance      =   0  'Flat
      Caption         =   "Negrita"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6960
      TabIndex        =   20
      Top             =   3960
      Width           =   975
   End
   Begin VB.CheckBox chkCursiva 
      Appearance      =   0  'Flat
      Caption         =   "Cursiva"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6000
      TabIndex        =   19
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   89522179
      CurrentDate     =   38555
   End
   Begin VB.TextBox txtFecha 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtUsuario 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtNota 
      Height          =   1755
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   6495
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   855
      Left            =   6840
      Picture         =   "frmCC_Comunicados.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   5760
      Picture         =   "frmCC_Comunicados.frx":1B9B4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   4800
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   89522179
      CurrentDate     =   38555
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Muestra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   3840
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
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
      TabIndex        =   14
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   4
      Left            =   4560
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vigencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   3
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Comunicados del área de servicios"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   6255
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
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
    txtCodigo = rs!cod_comunicado
    txtCodigo.Tag = 0
    txtCodigo.ToolTipText = rs!cod_comunicado
    
    txtUsuario.Text = rs!Usuario
    txtFecha.Text = rs!Fecha

    txtNota.Text = rs!nota
    dtpInicio.Value = rs!inicio
    dtpCorte.Value = rs!corte

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
          & " values('" & glogon.Usuario & "',getdate(),'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
          & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" & txtNota.Text & "','" & cboFuente.Text _
          & "','" & Mid(cboColor.Text, 1, 1) & "'," & chkCursiva.Value & "," & chkNegrita.Value & ")"
   glogon.Conection.Execute strSQL
   
   strSQL = "select max(cod_comunicado) as Ultimo from sif_comunicados"
   rs.Open strSQL, glogon.Conection, adOpenStatic
    xComunicado = rs!Ultimo
   rs.Close
   
Else
  strSQL = "update sif_comunicados set usuario = '" & glogon.Usuario & "',fecha = getdate(), inicio = '" _
          & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00',corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
          & " 23:59:59',nota = '" & txtNota.Text & "',fFuente = '" & cboFuente.Text & "', fColor = '" _
          & Mid(cboColor.Text, 1, 1) & "',fCursiva = " & chkCursiva.Value & ", fNegrita = " & chkNegrita.Value _
          & " Where cod_comunicado = " & txtCodigo.ToolTipText
  glogon.Conection.Execute strSQL
  
  xComunicado = txtCodigo.ToolTipText
  
End If

MsgBox "Comunicado # " & xComunicado & " Registrado Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

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
    
    rs.Open strSQL, glogon.Conection, adOpenStatic
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
  MsgBox Err.Description, vbCritical

End Sub


