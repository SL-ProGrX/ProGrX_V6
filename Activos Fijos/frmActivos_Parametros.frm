VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmActivos_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros Generales"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPermitirRegistroPeriodoCerrado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Permitir Registro de Activos en Periodos Cerrados (Solo para Migración)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   6732
   End
   Begin VB.CheckBox chkForzarCompras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Forzar el Registro en el Auxiliar de Activos Fijos, las compras de activos del periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   7335
   End
   Begin VB.CommandButton cmdInicio 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Establecer Mes Inicial"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2292
   End
   Begin VB.CheckBox chkForzarTipoActivo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Forzar la Utilización de los parámetros de Depreciación y Vida Util del Tipo de Activo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   7695
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CheckBox chkEnlaceSIFC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Integración con el Sistema Comercial y BackOffice"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.CheckBox chkEnlaceConta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Integración con el Sistema de Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   615
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
      Top             =   4080
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Guardar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Parametros.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3720
      TabIndex        =   14
      Top             =   3360
      Width           =   1332
      _Version        =   1441792
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo de Inicio del Módulo de Activos Fijos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   6972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Base de Cálculo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nombre de la Empresa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2052
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Contabilidad de Enlace"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2052
   End
End
Attribute VB_Name = "frmActivos_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
'
strSQL = "select coalesce(count(*),0) as Existe from Activos_parametros"
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe = 0 Then
  'Insertar
  strSQL = "insert Activos_parametros(cod_empresa,nombre_empresa,enlace_conta,enlace_sifc,Tipo_Anio" _
         & ",forzar_TipoActivo,RegistroCompras, REGISTRO_PERIODO_CERRADO) values(" _
         & cbo.ItemData(cbo.ListIndex) & ",'" & UCase(txtNombre) & "','" & IIf((chkEnlaceConta.Value = vbChecked), "S", "N") _
         & "','" & IIf((chkEnlaceSIFC.Value = vbChecked), "S", "N") & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
         & "'," & chkForzarTipoActivo.Value & "," & chkForzarCompras.Value _
         & "," & chkPermitirRegistroPeriodoCerrado.Value & ")"
Else
  'Actualizar
  strSQL = "update Activos_parametros set cod_empresa = " & cbo.ItemData(cbo.ListIndex) _
         & ",nombre_empresa = '" & UCase(txtNombre) _
         & "',enlace_conta = '" & IIf((chkEnlaceConta.Value = vbChecked), "S", "N") _
         & "',enlace_sifc = '" & IIf((chkEnlaceSIFC.Value = vbChecked), "S", "N") _
         & "',tipo_anio = '" & cboTipo.ItemData(cboTipo.ListIndex) _
         & "',forzar_TipoActivo = " & chkForzarTipoActivo.Value _
         & ",registroCompras = " & chkForzarCompras.Value _
         & ",REGISTRO_PERIODO_CERRADO = " & chkPermitirRegistroPeriodoCerrado.Value

End If
Call ConectionExecute(strSQL)
rs.Close

MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdInicio_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date

strSQL = "select * from Activos_parametros"
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
   MsgBox "No se han guardado los parámetros, debe guardarlos primero y luego establecer el" _
          & " inicio del módulo.", vbExclamation
Else
   strSQL = "update Activos_parametros set inicio_anio = " & Year(dtpInicio.Value) _
          & ",inicio_mes = " & Month(dtpInicio.Value)
   Call ConectionExecute(strSQL)
   
   dtpInicio.Enabled = False
   cmdInicio.Enabled = False
   
   vFecha = DateAdd("m", -1, dtpInicio.Value)
   
   strSQL = "insert Activos_periodos(anio,mes,estado,asientos,traslado) values(" _
          & Year(vFecha) & "," & Month(vFecha) & ",'C','G','G')"
   Call ConectionExecute(strSQL)
   
   MsgBox "Periodo Inicial, para el modulo de activos fijos establecido correctamente...", vbInformation
      
   
End If
rs.Close

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
vModulo = 36


cboTipo.AddItem "Año Base para Cálculos con 360 días"
cboTipo.ItemData(cboTipo.NewIndex) = 1
cboTipo.AddItem "Año Base para Cálculos con 365 días"
cboTipo.ItemData(cboTipo.NewIndex) = 5

dtpInicio.Value = fxFechaServidor

strSQL = "select cod_Contabilidad,nombre from CntX_Contabilidades"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cbo.AddItem rs!Nombre
 cbo.ItemData(cbo.NewIndex) = rs!COD_CONTABILIDAD
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cbo.Text = rs!Nombre
End If
rs.Close


strSQL = "select * from Activos_parametros"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
   chkEnlaceConta.Value = IIf(rs!Enlace_Conta = "S", vbChecked, vbUnchecked)
   chkEnlaceSIFC.Value = IIf(rs!Enlace_SIFC = "S", vbChecked, vbUnchecked)
   chkPermitirRegistroPeriodoCerrado.Value = rs!REGISTRO_PERIODO_CERRADO
   
   txtNombre = rs!nombre_empresa
   
   chkForzarTipoActivo.Value = rs!forzar_TipoActivo
   chkForzarCompras.Value = rs!registroCompras
   
   If rs!tipo_anio = "5" Then
        cboTipo.Text = "Año Base para Cálculos con 365 días"
   Else
        cboTipo.Text = "Año Base para Cálculos con 360 días"
   End If
   
   
   If Not IsNull(rs!inicio_anio) Then
      dtpInicio.Enabled = False
      cmdInicio.Enabled = False
   End If
   
   
   
   strSQL = "select nombre from CntX_Contabilidades where cod_Contabilidad = " & rs!cod_empresa
   rs.Close
   
   Call OpenRecordSet(rs, strSQL, 0)
   If Not rs.EOF And Not rs.BOF Then
      cbo.Text = rs!Nombre
   End If
End If
rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub
