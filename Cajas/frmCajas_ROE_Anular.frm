VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_ROE_Anular 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ROE: Anulación"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   14895
      _Version        =   1572864
      _ExtentX        =   26273
      _ExtentY        =   9551
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
   Begin XtremeSuiteControls.GroupBox gbAnular 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   14655
      _Version        =   1572864
      _ExtentX        =   25850
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Anulación de ROE's "
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAnular 
         Height          =   615
         Left            =   12240
         TabIndex        =   13
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Anular"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_ROE_Anular.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1050
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1852
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Notas de la Anulación"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10393
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkFecha 
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Index           =   0
      Left            =   8160
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Index           =   1
      Left            =   9480
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   11040
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_ROE_Anular.frx":05A4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   12360
      TabIndex        =   7
      ToolTipText     =   "Exportar a Excel"
      Top             =   1080
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_ROE_Anular.frx":0CA4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   11040
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   233
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta y Anulación de ROE's"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   5505
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmCajas_ROE_Anular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnAnular_Click()
On Error GoTo vError

Dim i As Long, iCasos As Long


If Len(txtNotas.Text) <= 30 Then
           MsgBox "No se ha especificado una nota válida para la anulación!", vbExclamation
           Exit Sub
End If

Me.MousePointer = vbHourglass

iCasos = 0


With lsw.ListItems
 
 For i = 1 To .Count
  If .Item(i).Checked Then
        strSQL = "exec spCajas_ROE_Anula " & .Item(i).Text & ", '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Pass = 0 Then
            
           Me.MousePointer = vbDefault
           
           MsgBox rs!Mensaje, vbCritical
           Call btnBuscar_Click
           Exit Sub
        End If
        
        iCasos = iCasos + 1
  End If
 Next i

End With
Me.MousePointer = vbDefault

If iCasos > 0 Then
    MsgBox "Anulación de ROE's  Procesada!", vbInformation
    Call btnBuscar_Click
Else
    MsgBox "Seleccione algún ROE para su anulación!", vbExclamation
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()
On Error GoTo vError


Dim vWhere As Boolean

Me.MousePointer = vbHourglass

vWhere = True

strSQL = "select ID_ROE, TIPOROE, rtrim(CEDULA_ASO) as 'CEDULA_ASO', IDENTIFICACION_DEPO, NOMBRE_DEPO, FECHA, USUARIO, MONTO_LOCAL, MONTO_DOL,TIPO_CAMBIO" _
      & ", REGISTRO_FECHA, REGISTRO_USUARIO, ACTUALIZA_FECHA, ACTUALIZA_USUARIO, USUARIO_ANULACION, FECHA_ANULACION, OBSERV_ANULACION, IMPRIME_FECHA, IMPRIME_USUARIO " _
      & " , ISNULL(ID_SESION,'') AS 'ID_SESION', ESTADO" _
      & " From CAJAS_ROE WHERE ESTADO = 'A'"
      
If chkFecha.Value = xtpUnchecked Then
   If vWhere Then
        strSQL = strSQL & " AND "
   Else
        strSQL = strSQL & " WHERE "
        vWhere = True
   End If
   
   strSQL = strSQL & " Fecha between '" & Format(dtpFecha(0).Value, "yyyy-mm-dd") & "' and '" & Format(dtpFecha(1).Value, "yyyy-mm-dd") & "'"
End If

If Len(txtCedula.Text) > 0 Then

    txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
    
       If vWhere Then
            strSQL = strSQL & " AND "
       Else
            strSQL = strSQL & " WHERE "
            vWhere = True
       End If
       
       strSQL = strSQL & " (Cedula_Aso like '%" & txtCedula.Text & "%' or  IDENTIFICACION_DEPO like '%" & txtCedula.Text & "%') "
End If

If Len(txtNombre.Text) > 0 Then

    txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)
    
       If vWhere Then
            strSQL = strSQL & " AND "
       Else
            strSQL = strSQL & " WHERE "
            vWhere = True
       End If
       
       strSQL = strSQL & " (NOMBRE_DEPO like '%" & txtNombre.Text & "%' or  NOMBRE_DEPO like '%" & txtNombre.Text & "%') "
End If


lsw.ListItems.Clear
'ID_ROE, TIPOROE, CEDULA_ASO, IDENTIFICACION_DEPO, NOMBRE_DEPO, FECHA, USUARIO, MONTO_LOCAL, MONTO_DOL,TIPO_CAMBIO
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!ID_ROE)
     itmX.SubItems(1) = IIf(rs!Estado = "A", "ACTIVO", "ANULADO")
     
     itmX.SubItems(2) = rs!TIPOROE
     itmX.SubItems(3) = rs!CEDULA_ASO
     itmX.SubItems(4) = rs!Identificacion_Depo & ""
     itmX.SubItems(5) = rs!Nombre_Depo & ""
     itmX.SubItems(6) = rs!fecha
     itmX.SubItems(7) = rs!Usuario

     itmX.SubItems(8) = Format(rs!Monto_Local, "Standard")
     itmX.SubItems(9) = Format(rs!Monto_Dol, "Standard")
     itmX.SubItems(10) = Format(rs!TIPO_CAMBIO, "Standard")


     itmX.SubItems(11) = rs!Registro_Fecha & ""
     itmX.SubItems(12) = rs!Registro_Usuario & ""
     itmX.SubItems(13) = rs!actualiza_fecha & ""
     itmX.SubItems(14) = rs!actualiza_Usuario & ""
     itmX.SubItems(15) = rs!Imprime_fecha & ""
     itmX.SubItems(16) = rs!Imprime_Usuario & ""

     itmX.SubItems(17) = rs!USUARIO_ANULACION & ""
     itmX.SubItems(18) = rs!FECHA_ANULACION & ""
     itmX.SubItems(19) = rs!OBSERV_ANULACION & ""

     itmX.SubItems(20) = rs!Id_Sesion & ""

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnImprimir_Click()
On Error GoTo vError

Dim i As Long, iCasos As Long

Me.MousePointer = vbHourglass


iCasos = 0
With lsw.ListItems
For i = 1 To .Count - 1
    If .Item(i).Checked Then
        Call sbCajas_ROE(.Item(i).Text)
        iCasos = iCasos + 1
    End If
Next i
End With

Me.MousePointer = vbDefault

If iCasos > 0 Then
        MsgBox "ROE's Emitidos Satisfactoriamente!", vbInformation
Else
    MsgBox "Seleccione de la Lista los ROE que desea reimprimir!", vbExclamation
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub chkFecha_Click()
If chkFecha.Value = xtpChecked Then
    dtpFecha(0).Enabled = False
Else
    dtpFecha(0).Enabled = True
End If

dtpFecha(1).Enabled = dtpFecha(0).Enabled
End Sub

Private Sub Form_Load()

vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1500, vbCenter
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Tipo ROE", 1500, vbCenter
    .Add , , "Cliente Id", 2100, vbCenter
    .Add , , "Depositante Id", 2100, vbCenter
    .Add , , "Nombre Depositante", 3500
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Usuario", 2500
    .Add , , "Monto Local", 2100, vbRightJustify
    .Add , , "Monto Dólares", 2100, vbRightJustify
    .Add , , "Tipo Cambio", 1100, vbRightJustify

    .Add , , "Registro Fecha", 2500, vbCenter
    .Add , , "Registro Usuario", 2500
    .Add , , "Actualiza Fecha", 2500, vbCenter
    .Add , , "Actualiza Usuario", 2500
    .Add , , "Imprime Fecha", 2500, vbCenter
    .Add , , "Imprime Usuario", 2500

    .Add , , "Anula Fecha", 2500, vbCenter
    .Add , , "Anula Usuario", 2500
    .Add , , "Anula Nota", 3500

    .Add , , "Sesión Id", 1200, vbCenter


End With


dtpFecha(0).Value = fxFechaServidor
dtpFecha(1).Value = dtpFecha(0).Value

Call chkFecha_Click

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

GLOBALES.gTag = "ROE_" & Item.Text
Call sbFormsCall("frmCajas_ROE", vbModal, , , False, Me, True)

End Sub


