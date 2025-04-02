VERSION 5.00
Begin VB.Form frmAF_Renuncia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renuncia"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   HelpContextID   =   1010
   Icon            =   "frmAF_Renuncia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkcausamortalidad 
      Alignment       =   1  'Right Justify
      Caption         =   "Renuncia por Mortalidad"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton optRenunciaPat 
      Caption         =   "Renuncia al Patrono"
      Height          =   315
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton optRenunciaAso 
      Caption         =   "Renuncia a la Asociación"
      Height          =   315
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox chkningunacausa 
      Alignment       =   1  'Right Justify
      Caption         =   "Ninguna causa de renuncia "
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar "
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdPlantear 
      Caption         =   "&Aplicar"
      Height          =   375
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cboCausaRenuncia 
      Height          =   315
      ItemData        =   "frmAF_Renuncia.frx":08CA
      Left            =   720
      List            =   "frmAF_Renuncia.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aplicación de la Renuncia del Socio"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Causa"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmAF_Renuncia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrecSocio As New ADODB.Recordset
Sub ActualizaSocio()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      De acuerdo al tipo de renuncia aplicada, actualiza en la tabla socios el
'               nuevo estado del Socio.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo ErrorTransaccion
glogon.Conection.BeginTrans

If optRenunciaAso(0).Value = True Then
   glogon.Conection.Execute "Update Socios SET EstadoActual='A',Ultimo_Estado='S',ind_liquidacion=1 WHERE Cedula = '" & mrecSocio!Cedula & "'"
   Call Bitacora("Modifica", "Modifico estado del Socio " & Trim(mrecSocio!Cedula))
Else
   glogon.Conection.Execute "Update Socios SET EstadoActual='P',Ultimo_Estado='A',ind_liquidacion=1 WHERE Cedula = '" & mrecSocio!Cedula & "'"
   Call Bitacora("Modifica", "Modifico estado del Socio " & Trim(mrecSocio!Cedula))
End If

glogon.Conection.CommitTrans
    
Exit Sub
ErrorTransaccion:
glogon.Conection.RollbackTrans
  Call ProcedimientoErrores(Me.Name, Err)
  
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdPlantear_Click(Index As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Guarda en la Tabla renuncias la causa por la cual esta renunciando el socio.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               ActualizaSocio - (Actualiza el nuevo estado del socio)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'               fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim strFecha As String, strCedula As String

On Error GoTo Errores
Me.MousePointer = vbHourglass

If optRenunciaAso(0).Value = False And optRenunciaPat(1).Value = False Then
   MsgBox "No Puede Plantear La Renuncia", vbExclamation, "Faltan Datos"
ElseIf Trim(cboCausaRenuncia) = "" And chkningunacausa(0).Value = vbUnchecked And chkcausamortalidad(1).Value = vbUnchecked Then
   MsgBox "No Puede Plantear La Renuncia", vbExclamation, "Faltan Datos"
Else
    strCedula = Trim(frmAF_Principal.txtCedula)
    mrecSocio.Source = "Select Cedula,Nacta,id_Promotor,Id_Boleta_Af,Ind_Liquidacion from Socios Where Cedula='" & Trim(frmAF_Principal.txtCedula) & "'"
    mrecSocio.ActiveConnection = glogon.Conection
    mrecSocio.CursorType = adOpenStatic
    mrecSocio.Open
    
    If mrecSocio!ind_Liquidacion = 1 Then
       MsgBox "No Puede Plantear La Renuncia", vbExclamation, "Antes Efectue La Liquidacion A La Asociación"
       Me.MousePointer = vbDefault
       mrecSocio.Close
       Exit Sub
    End If

'------------------------------------------------
  strSQL = "Insert into Renuncias("
  If chkningunacausa(0).Value = vbUnchecked And chkcausamortalidad(1).Value = vbUnchecked Then
     strSQL = strSQL & "ID_Causa,"
  End If
  strSQL = strSQL & "ID_Promotor,Cedula,Id_Boleta,"
  
  If optRenunciaAso(0).Value = True Then
     strSQL = strSQL & "FechaRenA,"
  Else
     strSQL = strSQL & "FechaRenP,"
  End If
  strSQL = strSQL & "TipoRen,Nacta"
  
  If chkningunacausa(0).Value = vbChecked Then
     strSQL = strSQL & ",NCausaRen"
  ElseIf chkcausamortalidad(1).Value = vbChecked Then
     strSQL = strSQL & ",RenMor"
  End If
  
  strSQL = strSQL & ") Values("
    
  If chkningunacausa(0).Value = vbUnchecked And chkcausamortalidad(1).Value = vbUnchecked Then
    rs.Source = "Select Id_Causa from Causas_Renuncias Where Descripcion='" & Trim(cboCausaRenuncia) & "'"
    rs.Open , glogon.Conection, adOpenStatic

    If rs.EOF = False Then
       strSQL = strSQL & CLng(rs!Id_Causa) & ","
    End If
    rs.Close
  End If
'------------------------------------------------
    strSQL = strSQL & CLng(mrecSocio!Id_Promotor) & ",'"
    strSQL = strSQL & CStr(Trim(mrecSocio!Cedula)) & "',"
    strSQL = strSQL & CLng(mrecSocio!Id_Boleta_Af) & ",'"
    
    strFecha = Format(fxFechaServidor, "yyyy/mm/dd")
    If optRenunciaAso(0).Value = True Then
      strSQL = strSQL & strFecha & "',"
      strSQL = strSQL & "'A',"
    Else
      strSQL = strSQL & strFecha & "',"
      strSQL = strSQL & "'P',"
    End If
    
    strSQL = strSQL & CLng(IIf(IsNull(mrecSocio!Nacta), 0, mrecSocio!Nacta))
    
    If chkningunacausa(0).Value = vbChecked Then
       strSQL = strSQL & ",1"
    ElseIf chkcausamortalidad(1).Value = vbChecked Then
       strSQL = strSQL & ",1"
    End If
    
    strSQL = strSQL & ")"
    glogon.Conection.Execute (strSQL)
    
'------------------------------------------------
    Call Bitacora("Registra", "Registro renuncia al Socio " & mrecSocio!Cedula)
'------------------------------------------------
    Call ActualizaSocio
    mrecSocio.Close
    MsgBox "Renuncia Aplicada", vbExclamation, "Registro Actualizado"
    frmAF_Principal.txtCedula = ""
    GLOBALES.gblnBuscando = True
    frmAF_Principal.txtCedula = strCedula
    Unload Me
End If

Me.MousePointer = vbDefault

Exit Sub
Errores:
    Me.MousePointer = vbDefault
    Call ProcedimientoErrores(Me.Name, Err)
End Sub

Private Sub chkcausamortalidad_Click(Index As Integer)
  chkningunacausa(0).Value = vbUnchecked
End Sub

Private Sub chkningunacausa_Click(Index As Integer)
 chkcausamortalidad(1).Value = vbUnchecked
End Sub


Private Sub Form_DblClick()
'Set Conlsw.frmX = Me
'Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Establecer permisos sobre el formulario, Carga lista de Causas de renuncia.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim rs As New ADODB.Recordset

On Error GoTo CapturaError

vModulo = 1
Call Formularios(Me)
Call RefrescaTags(Me)

If Trim(frmAF_Principal.txtEstadoActual) = "Renunció a la Asociacion" Then
   optRenunciaAso(0).Enabled = False
   optRenunciaPat(1).Value = True
ElseIf Trim(frmAF_Principal.txtEstadoActual) = "Socio" Then
   optRenunciaPat(1).Enabled = False
   optRenunciaAso(0).Value = True
End If

rs.Open "select Descripcion from causas_renuncias", glogon.Conection, adOpenStatic
Do While Not rs.EOF
  cboCausaRenuncia.AddItem rs!Descripcion
  rs.MoveNext
Loop
rs.Close

Exit Sub
CapturaError:
   Call ProcedimientoErrores(Me.Name, Err)
End Sub

