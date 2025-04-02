VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_Plan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mesanjes del Comité"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   Icon            =   "frmAF_CD_Plan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAF_CD_Plan.frx":3482
   ScaleHeight     =   7095
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMsj 
      Caption         =   "Crear nuevo mensaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtMsj 
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
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1200
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker dtpMsjVence 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   169672705
         CurrentDate     =   37679
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmAF_CD_Plan.frx":378E
         Top             =   360
         Width           =   720
      End
      Begin VB.Image imgGuardaMsj 
         Height          =   255
         Left            =   4560
         Picture         =   "frmAF_CD_Plan.frx":3AC0
         Stretch         =   -1  'True
         ToolTipText     =   "Guardar Mensaje"
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha de vencimiento del Mensaje "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Image imgMsjCierraFrame 
         Height          =   255
         Left            =   4920
         Picture         =   "frmAF_CD_Plan.frx":3BAC
         Stretch         =   -1  'True
         ToolTipText     =   "No Incluir Mensaje"
         Top             =   3240
         Width           =   255
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5865
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   7755
      _Version        =   524288
      _ExtentX        =   13679
      _ExtentY        =   10345
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CD_Plan.frx":3CD3
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label txtComite 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblComite 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Comité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgBorraMsj 
      Height          =   240
      Left            =   7560
      Picture         =   "frmAF_CD_Plan.frx":4273
      ToolTipText     =   "Borrar mensaje"
      Top             =   480
      Width           =   240
   End
   Begin VB.Image imgMsjNuevo 
      Height          =   240
      Left            =   7200
      Picture         =   "frmAF_CD_Plan.frx":AAC5
      ToolTipText     =   "Nuevo mensaje"
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "frmAF_CD_Plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Sub sbVerificaComite()
  strSQL = "select top 1 cod_comite, DESCRIPCION from afi_cd_comites where cod_comite = '" & txtComite.Caption & "'"
           rs.Open strSQL, glogon.Conection, adOpenStatic
       If rs.EOF Then
          MsgBox "Este Comité no esta registrado", vbCritical, "Verificación"
       Else
          lblComite.Caption = Trim(rs!Descripcion)
       End If
       rs.Close
 End Sub
Private Function fxNumMensaje() As Long

Dim strSQL As String, rs As New ADODB.Recordset

    strSQL = "Select coalesce(Max(num_mensaje),0) as Con from afi_cd_comites_mensajes"
    rs.Open strSQL, glogon.Conection, adOpenStatic
      fxNumMensaje = (rs!Con + 1)
    rs.Close

End Function
Private Sub sbCargaMsj(vCodigo As String)

Dim strSQL As String, rs As New ADODB.Recordset

dtpMsjVence = fxFechaServidor + 30

On Error GoTo vError

Me.MousePointer = vbHourglass

'Inicializa Datos y Encabezados
dtpMsjVence.Value = fxFechaServidor
vGrid.MaxRows = 0
vGrid.MaxCols = 3

txtMsj = ""
fraMsj.Visible = False

strSQL = "select * from afi_cd_comites_mensajes where cod_comite = '" _
       & vCodigo & "' and datediff(d,getdate(),vencimiento) >= 0"
       rs.Open strSQL, glogon.Conection, adOpenForwardOnly

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  vGrid.Text = Format(rs!vencimiento, "dd/mm/yyyy")
  vGrid.TextTip = TextTipFixed
  vGrid.TextTipDelay = 1000

  vGrid.CellNote = "Fecha : " & rs!Fecha & vbCrLf & "Usuario : " & rs!Usuario
  vGrid.CellTag = rs!Usuario
   
  vGrid.Col = 2
  vGrid.Text = rs!Mensaje
  
  vGrid.RowHeight(vGrid.Row) = vGrid.MaxTextRowHeight(vGrid.Row)
  
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub
Private Sub Form_Load()
  
 vModulo = 23
  
  dtpMsjVence = fxFechaServidor + 365
  If GLOBALES.gTag <> Empty Then
    txtComite.Caption = GLOBALES.gTag
    TxtComite_KeyPress (vbKeyReturn)
  End If
  GLOBALES.gTag = Empty

End Sub

Private Sub imgBorraMsj_Click()
Dim strSQL As String, i As Integer
Dim msj(2) As String

On Error GoTo vError

With vGrid
    For i = 1 To vGrid.MaxRows
      .Row = i
      .Col = 3
        If .Value = 1 Then
           .Col = 1
           msj(0) = Format(.Text, "yyyymmdd")
           msj(1) = .CellTag
           .Col = 2
           msj(2) = Mid(.Text, 1, 15)
           
           strSQL = "delete afi_cd_comites_mensajes where cod_comite = '" & txtComite.Caption _
                  & "' and usuario = '" & msj(1) & "' and vencimiento = '" _
                  & msj(0) & "' and substring(mensaje,1,15) = '" _
                  & msj(2) & "'"
           glogon.Conection.Execute strSQL
        End If
    Next i
End With

Call sbCargaMsj(txtComite.Caption)

Exit Sub

vError:
 MsgBox Err.Description, vbExclamation
End Sub

Private Sub imgGuardaMsj_Click()
    Dim strSQL As String
    
    
    On Error GoTo vError
    
    strSQL = "insert afi_cd_comites_mensajes(fecha,usuario,cod_comite,vencimiento,mensaje,num_mensaje) " _
           & "values(getdate(),'" & glogon.Usuario & "','" & txtComite.Caption & "','" & Format(dtpMsjVence.Value, "yyyymmdd") & "','" _
           & txtMsj & "'," & fxNumMensaje & ")"
           glogon.Conection.Execute strSQL
    
    txtMsj = ""
    fraMsj.Visible = False
    MsgBox "Mensaje Registrado...", vbInformation
    
    Call sbCargaMsj(txtComite.Caption)
    
    
    Exit Sub
vError:
     MsgBox Err.Description, vbCritical
End Sub

Private Sub imgMsjCierraFrame_Click()
  fraMsj.Visible = False
End Sub

Private Sub imgMsjNuevo_Click()
  
  fraMsj.Visible = True
  dtpMsjVence = fxFechaServidor + 365
End Sub


Private Sub TxtComite_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
     Call sbVerificaComite
     Call sbCargaMsj(txtComite.Caption)
    Case Else
    KeyAscii = 0
End Select

End Sub

