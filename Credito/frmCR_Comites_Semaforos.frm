VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Comites_Semaforos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Semáforos para Resolución de Expedientes de Crédito"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   3836
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.PushButton btnSemaforo 
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   600
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Guardar"
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
      Appearance      =   21
      Picture         =   "frmCR_Comites_Semaforos.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   4440
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13996
      _ExtentY        =   661
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboComites 
      Height          =   330
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   330
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.FlatEdit txtRojo 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Text            =   "5"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAmarillo 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   840
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Text            =   "5"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnEmail 
      Height          =   615
      Left            =   7200
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Agregar"
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
      Appearance      =   21
      Picture         =   "frmCR_Comites_Semaforos.frx":0731
   End
   Begin XtremeSuiteControls.PushButton btnEmail_Delete 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   1800
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
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
      Appearance      =   21
      Picture         =   "frmCR_Comites_Semaforos.frx":0E51
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Correos para Notificación de Resoluciones:"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Email: "
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   5280
      Picture         =   "frmCR_Comites_Semaforos.frx":13F5
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   5280
      Picture         =   "frmCR_Comites_Semaforos.frx":1A03
      Top             =   360
      Width           =   240
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Unidad"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Comité"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCR_Comites_Semaforos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnEmail_Click()
If vPaso Then Exit Sub

On Error GoTo vError

If fxEmail_Valida(txtEmail.Text) Then
     '(@ComiteId int, @Email varchar(100), @Usuario varchar(30))
   strSQL = "exec spCrd_Comites_Semaforo_Email_Add " & cboComites.ItemData(cboComites.ListIndex) _
          & ", '" & txtEmail.Text & "', '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
   If rs!Pass = 1 Then
      Call Bitacora(rs!Movimiento, rs!Detalle)
      MsgBox "Correo Registrado Satisfactoriamente! ", vbInformation
   Else
     MsgBox "No fue Posible registrar el Correo, verifique! " & rs!Mensaje, vbExclamation
   End If
 
End If

Call cboComites_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub btnEmail_Delete_Click()

If vPaso Then Exit Sub

Dim i As Integer, iCasos As Integer

On Error GoTo vError

iCasos = 0
With lsw.ListItems
    For i = 1 To .Count
        If .Item(i).Checked Then
            iCasos = iCasos + 1
        End If
    Next i

    If iCasos = 0 Then
        MsgBox "Seleccione los correos que desea eliminar!", vbExclamation
        Exit Sub
    End If

    For i = 1 To .Count
        If .Item(i).Checked Then
        
            strSQL = "exec spCrd_Comites_Semaforo_Email_Delete " & .Item(i).Text _
                   & ", '" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Comite [" & cboComites.ItemData(cboComites.ListIndex) & "] Email: " _
                    & .Item(i).SubItems(1))
            
        End If
    Next i
End With

Call cboComites_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnSemaforo_Click()

If vPaso Then Exit Sub

On Error GoTo vError
'spCrd_Comites_Semaforo_Add(@ComiteId int, @UnidadTiempo varchar(30), @AlertaRoja int, @AlertaAmarilla int, @Usuario varchar(30))
Dim pUnidad As String

Select Case Mid(cboUnidad, 1, 1)
    Case "D"
        pUnidad = "DAY"
    Case "M"
        pUnidad = "MINUTE"
    Case "H"
        pUnidad = "HOUR"
End Select


strSQL = "exec spCrd_Comites_Semaforo_Add " & cboComites.ItemData(cboComites.ListIndex) _
       & ", '" & pUnidad & "', " & txtRojo.Text & ", " & txtAmarillo.Text & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then
   Call Bitacora(rs!Movimiento, rs!Detalle)
   MsgBox "Semaforo Registrado Satisfactoriamente! ", vbInformation
Else
  MsgBox "No fue Posible registrar el Semaforo, verifique! " & rs!Mensaje, vbExclamation
End If
 
Call cboComites_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboComites_Click()

If vPaso Then Exit Sub

txtAmarillo.Text = "0"
txtRojo.Text = "0"
txtEmail.Text = ""

strSQL = "SELECT *" _
       & "     , case when UnidadTiempo = 'MINUTE' Then 'Minutos'" _
       & "    when UnidadTiempo = 'DAY'    Then 'Días'" _
       & "    when UnidadTiempo = 'HOUR'   Then 'Horas'" _
       & "    Else '' end 'UnidadTiempoEsp' " _
       & " FROM CRD_COMITES_SEMAFORO WHERE IdComite = " & cboComites.ItemData(cboComites.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtRojo.Text = rs!AlertaRoja
   txtAmarillo.Text = rs!AlertaAmarilla
   
   cboUnidad.Text = rs!UnidadTiempoEsp
End If

strSQL = "select * FROM CRD_COMITES_SEMAFORO_EMAIL WHERE IdComite = " & cboComites.ItemData(cboComites.ListIndex)
lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!IdRegistro)
     itmX.SubItems(1) = rs!Email & ""
     itmX.SubItems(2) = rs!UsuarioInserta & ""
     itmX.SubItems(3) = rs!FechaInserta & ""
 rs.MoveNext
Loop
rs.Close


End Sub

Private Sub Form_Load()

vModulo = 3

cboUnidad.AddItem "Minutos"
cboUnidad.AddItem "Horas"
cboUnidad.AddItem "Días"
cboUnidad.Text = "Días"

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Email", 3000
    .Add , , "Usuario", 2100
    .Add , , "Fecha", 2100
End With
lsw.Checkboxes = True


vPaso = True
    strSQL = "select Id_Comite as 'Idx', Descripcion as 'ItmX'" _
           & "  from COMITES   Where estado = 1"
    Call sbCbo_Llena_New(cboComites, strSQL, False, True)
vPaso = False

Call cboComites_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub
