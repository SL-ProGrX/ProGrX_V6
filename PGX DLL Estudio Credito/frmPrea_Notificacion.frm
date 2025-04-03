VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPrea_Notificacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Envío de Notificación"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTiquete 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1560
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2400
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9763
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCelular 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2040
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.PushButton btnNotificar 
      Height          =   450
      Left            =   7080
      TabIndex        =   11
      ToolTipText     =   "Copiar Expediente"
      Top             =   120
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Notificar"
      BackColor       =   16761024
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
      Picture         =   "frmPrea_Notificacion.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   615
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16960
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "Enviar Notificación de Resolución"
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
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Móvil"
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
      Height          =   330
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Correo"
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
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tiquete"
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
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Sugerido"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrea_Notificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNotificar_Click()



'spCRD_PREA_CONSULTA_PLANTILLA_NOTIFICACION

'spCRD_PREA_CONSULTA_PLANTILLA_MENSAJE
'spCRD_PREA_CONSULTA_PLANTILLA_MENSAJE_SMS
'
'If Trim(txtEstado.Text) = "Denegado" Then
'                TipoPlantilla = "DENEGA"
'                vEstado = "DESC"
'                DatosPlantilla.NombreAsociado = nomAsociado
'            ElseIf Trim(txtEstado.Text) = "Aprobado" And Convert.ToDecimal(txtMontoSugerido.Text) > 0 Then
'                TipoPlantilla = "APROCO"
'                vEstado = "APRO"
'                DatosPlantilla.MontoAprobado = Format(MontoAprobado, "Standard")
'                DatosPlantilla.MontoSugerido = Format(txtMontoSugerido.Text, "Standard")
'                DatosPlantilla.NombreAsociado = nomAsociado
'                DatosPlantilla.Tiquete = txtTiquete.Text
'            ElseIf Trim(txtEstado.Text) = "Aprobado" And Convert.ToDecimal(txtMontoSugerido.Text) = 0 Then
'                TipoPlantilla = "APROSO"
'                vEstado = "APRO"
'                DatosPlantilla.MontoAprobado = Format(MontoAprobado, "Standard")
'                DatosPlantilla.NombreAsociado = nomAsociado
'                DatosPlantilla.Tiquete = txtTiquete.Text
'            End If
'
'
'            lPlantilla = notificacionesBLL.ObtienePlantillaCorreo(TipoPlantilla)
'
'            For Each plantilla As PlantillaCorreo In lPlantilla
'                Asunto = plantilla.Asunto
'                Cuerpo = plantilla.plantilla
'            Next
'
'            lPlantillaMensaje = notificacionesBLL.ObtienePlantillaMensaje(TipoPlantilla)
'            lPlantillaMensajeSMS = notificacionesBLL.ObtienePlantillaMensajeSMS(TipoPlantilla)
'
'            For Each plantilla As PlantillaCorreo In lPlantillaMensajeSMS
'                Mensaje_sms = plantilla.Mensaje_sms
'
'            Next
'            For Each plantilla As PlantillaCorreo In lPlantillaMensaje
'                Mensaje = plantilla.Mensaje
'
'            Next
'
'            PlantillaCompleta = notificacionesBLL.CompletarPlantilla(Cuerpo, DatosPlantilla)
'            PlantillaCompletaMSJ = notificacionesBLL.CompletarPlantillaMensaje(Mensaje, DatosPlantilla)
'
'


'Select dbo.fxCrdPreaValidaUsuarioEnviaNotificacion(@Usuario,@Estado) as Resultado
'vUsuarioNotifica = notificacionesBLL.fxValidaUsuarioEnviaNotificacion(glogon.Usuario, vEstado)
'
'
'            If vUsuarioNotifica = "NMYC" Then 'notifica mensaje y correo
'                If txtCorreo.Text.Length > 0 And txtCelular.Text.Length > 0 Then
'                    ''Guardar en bitacora el nuevo registro
'                    oAuditoria = oUtilitario.GenerarBitacoraSIF(vModulo)

'spCRD_PREA_NOTIFICA_ENVIA_ALERTA

'                    notificacionesBLL.RegistraNotificacionCorreoYMsjTexto(PlantillaCompleta, Asunto, "P", Trim(txtCorreo.Text), "ECO", glogon.Usuario, "")
'                    notificacionesBLL.RegistraNotificacionCorreoYMsjTexto(Mensaje_SMS, Asunto, "P", Trim(txtCorreo.Text), "EMJ", glogon.Usuario, Trim(txtCelular.Text))
'                    notificacionesBLL.RegistraBitacoraCorreo(codPreanalisis, PlantillaCompletaMSJ, Mensaje_SMS, txtCorreo.Text, txtCelular.Text, "SI", "SI", glogon.Usuario)
'                    'REGISTRO EN BITACORA
'                    oAuditoria.Detalle = "Se envia notificación de resolución de estudio de crédito: " & " Preanalisis: " & codPreanalisis & " Cedula: " & cedulaAsoc & " Asunto: " & Asunto & " Correo: " & Trim(txtCorreo.Text) & " Teléfono: " & Trim(txtCelular.Text) & " Usuario envía:" & glogon.Usuario
'                    oAuditoria.Origen = "Este evento se generó desde la clase: " & MethodBase.GetCurrentMethod().DeclaringType.Name & " dentro del método: " & MethodBase.GetCurrentMethod().Name & " en la línea: " & New StackFrame(0, True).GetFileLineNumber()
'                    oAuditoria.Movimiento = "REGISTRA"
'                    oAuditoria.Evento = 79
'                    'Se realiza la inserción en bitácora
'                    oBitacora.InsertaBitacoraSIF (oAuditoria)
'                    mMensajes.fxMensajeInformacion ("Se ha realizado la acción correctamente.")
'                    Close()
'                End If
'            ElseIf vUsuarioNotifica = "NSMJ" Then 'notifica solo msj
'                If txtCelular.Text.Length > 0 Then '  Sol 33976 mchaves
'                    ''Guardar en bitacora el nuevo registro
'                    oAuditoria = oUtilitario.GenerarBitacoraSIF(vModulo)
'                    notificacionesBLL.RegistraNotificacionCorreoYMsjTexto(Mensaje_SMS, Asunto, "P", Trim(txtCorreo.Text), "EMJ", glogon.Usuario, Trim(txtCelular.Text))
'                    notificacionesBLL.RegistraBitacoraCorreo(codPreanalisis, "", Mensaje_SMS, txtCorreo.Text, txtCelular.Text, "SI", "NO", glogon.Usuario) 'Sol:33976 mchaves
'                    'REGISTRO EN BITACORA
'                    oAuditoria.Detalle = "Se envia notificación de resolución de estudio de crédito: " & " Preanalisis: " & codPreanalisis & " Cedula: " & cedulaAsoc & " Asunto: " & Asunto & " Correo: " & Trim(txtCorreo.Text) & " Teléfono: " & Trim(txtCelular.Text) & " Usuario envía: " & glogon.Usuario
'                    oAuditoria.Origen = "Este evento se generó desde la clase: " & MethodBase.GetCurrentMethod().DeclaringType.Name & " dentro del método: " & MethodBase.GetCurrentMethod().Name & " en la línea: " & New StackFrame(0, True).GetFileLineNumber()
'                    oAuditoria.Movimiento = "REGISTRA"
'                    oAuditoria.Evento = 79
'                    'Se realiza la inserción en bitácora
'                    oBitacora.InsertaBitacoraSIF (oAuditoria)
'                    mMensajes.fxMensajeInformacion ("Se ha realizado la acción correctamente.")
'                    Close()
'                End If
'
'            ElseIf vUsuarioNotifica = "NSCO" Then 'notifica solo correo
'                If txtCorreo.Text.Length > 0 Then '  Sol 33976 mchaves
'                    ''Guardar en bitacora el nuevo registro
'                    oAuditoria = oUtilitario.GenerarBitacoraSIF(vModulo)
'                    notificacionesBLL.RegistraNotificacionCorreoYMsjTexto(PlantillaCompleta, Asunto, "P", Trim(txtCorreo.Text), "ECO", glogon.Usuario, "")
'                    notificacionesBLL.RegistraBitacoraCorreo(codPreanalisis, PlantillaCompletaMSJ, "", txtCorreo.Text, txtCelular.Text, "NO", "SI", glogon.Usuario) 'Sol:33976 mchaves
'                    'REGISTRO EN BITACORA
'                    oAuditoria.Detalle = "Se envia notificación de resolución de estudio de crédito: " & " Preanalisis: " & codPreanalisis & " Cedula: " & cedulaAsoc & " Asunto: " & Asunto & " Correo: " & Trim(txtCorreo.Text) & " Teléfono: " & Trim(txtCelular.Text) & " Usuario envía: " & glogon.Usuario
'                    oAuditoria.Origen = "Este evento se generó desde la clase: " & MethodBase.GetCurrentMethod().DeclaringType.Name & " dentro del método: " & MethodBase.GetCurrentMethod().Name & " en la línea: " & New StackFrame(0, True).GetFileLineNumber()
'                    oAuditoria.Movimiento = "REGISTRA"
'                    oAuditoria.Evento = 79
'                    'Se realiza la inserción en bitácora


'spBIT_InsertarBitacora

'                    oBitacora.InsertaBitacoraSIF (oAuditoria)
'
'                    mMensajes.fxMensajeInformacion ("Se ha realizado la acción correctamente.")
'                    Close()
'                End If
'
'            ElseIf vUsuarioNotifica = "NOAP" Then ' no notifica
'                mMensajes.fxMensajeInformacion ("Favor validar la configuración de mensajes, correo electrónico o teléfono registrado.")
'                Exit Sub
'            ElseIf vUsuarioNotifica = "OFNC" Then ' no hay oficina configurada  Sol 33976 mchaves
'                mMensajes.fxMensajeInformacion ("La oficina a la que pertenece el usuario, no esta autorizada para el envío de notificaciones, favor validar.")
'                Exit Sub
'            End If



End Sub
