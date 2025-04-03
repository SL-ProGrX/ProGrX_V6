VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_Plan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensajes del Comité"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   Icon            =   "frmAF_CD_Plan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   1920
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbMsj 
      Height          =   3975
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   7011
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.DateTimePicker dtpMsjVence 
         Height          =   330
         Left            =   3120
         TabIndex        =   4
         Top             =   3360
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtMsj 
         Height          =   2535
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
         _ExtentY        =   4471
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de vencimiento del Mensaje "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgMsjCierraFrame 
         Height          =   255
         Left            =   5520
         Picture         =   "frmAF_CD_Plan.frx":3482
         Stretch         =   -1  'True
         ToolTipText     =   "No Incluir Mensaje"
         Top             =   3360
         Width           =   255
      End
      Begin VB.Image imgGuardaMsj 
         Height          =   255
         Left            =   5160
         Picture         =   "frmAF_CD_Plan.frx":3B88
         Stretch         =   -1  'True
         ToolTipText     =   "Guardar Mensaje"
         Top             =   3360
         Width           =   255
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6135
         _Version        =   1441793
         _ExtentX        =   10821
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Crear Nuevo Mensaje"
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
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5865
      Left            =   360
      TabIndex        =   0
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CD_Plan.frx":4298
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtComite 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   7815
      _Version        =   1441793
      _ExtentX        =   13785
      _ExtentY        =   661
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comité"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image imgBorraMsj 
      Height          =   240
      Left            =   7440
      Picture         =   "frmAF_CD_Plan.frx":4882
      ToolTipText     =   "Borrar mensaje"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgMsjNuevo 
      Height          =   240
      Left            =   7080
      Picture         =   "frmAF_CD_Plan.frx":4E16
      ToolTipText     =   "Nuevo mensaje"
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmAF_CD_Plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbVerificaComite()

strSQL = "select cod_comite, DESCRIPCION from afi_cd_comites where cod_comite = '" & txtComite.Tag & "'"
Call OpenRecordSet(rs, strSQL)
       
If rs.EOF Then
   MsgBox "Este Comité no esta registrado, Consulte uno válido!", vbCritical
Else
   txtComite.Text = Trim(rs!Descripcion)
End If
rs.Close
 
 End Sub

Private Function fxNumMensaje() As Long
Dim pSQL As String
   pSQL = "Select coalesce(Max(num_mensaje),0) as Con from afi_cd_comites_mensajes"
    Call OpenRecordSet(rs, pSQL)
      fxNumMensaje = (rs!Con + 1)
    rs.Close

End Function

Private Sub sbCargaMsj(vCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

'Inicializa Datos y Encabezados
dtpMsjVence.Value = DateAdd("y", 5, fxFechaServidor)
txtMsj = ""

vGrid.MaxRows = 0
vGrid.MaxCols = 3

vGrid.Visible = True
gbMsj.Visible = False

strSQL = "select * from afi_cd_comites_mensajes" _
       & " where cod_comite = '" & vCodigo & "' and datediff(d,getdate(),vencimiento) >= 0"

Call OpenRecordSet(rs, strSQL)

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
  vGrid.CellTag = rs!num_Mensaje
  
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
  
 vModulo = 40
  
 txtComite.Tag = GLOBALES.gTag
  
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub imgBorraMsj_Click()
Dim i As Integer, pMsjNum As Long

On Error GoTo vError

With vGrid
    For i = 1 To vGrid.MaxRows
      .Row = i
      .Col = 3
        If .Value = vbChecked Then
           .Col = 2
            pMsjNum = .CellTag
            
           strSQL = "delete afi_cd_comites_mensajes where cod_comite = '" & txtComite.Tag _
                  & "' and num_mensaje = " & pMsjNum
           Call ConectionExecute(strSQL)
        End If
    Next i
End With

Call sbCargaMsj(txtComite.Tag)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
End Sub

Private Sub imgGuardaMsj_Click()

    
On Error GoTo vError

txtMsj.Text = fxSysCleanTxtInject(txtMsj.Text)

strSQL = "insert afi_cd_comites_mensajes(fecha,usuario,cod_comite,vencimiento,mensaje,num_mensaje) " _
       & "values(getdate(),'" & glogon.Usuario & "','" & txtComite.Tag & "','" & Format(dtpMsjVence.Value, "yyyy-mm-dd") & "','" _
       & txtMsj.Text & "'," & fxNumMensaje & ")"
Call ConectionExecute(strSQL)

MsgBox "Mensaje registrado satisfactoriamente!", vbInformation

Call sbCargaMsj(txtComite.Tag)

Call imgMsjCierraFrame_Click

Exit Sub

vError:
     MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgMsjCierraFrame_Click()
  vGrid.Visible = True
  gbMsj.Visible = False
End Sub

Private Sub imgMsjNuevo_Click()
  vGrid.Visible = False
  gbMsj.Visible = True
  dtpMsjVence.Value = DateAdd("y", 5, fxFechaServidor)
  txtMsj.Text = ""
  txtMsj.SetFocus
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbVerificaComite
Call sbCargaMsj(txtComite.Tag)

End Sub


