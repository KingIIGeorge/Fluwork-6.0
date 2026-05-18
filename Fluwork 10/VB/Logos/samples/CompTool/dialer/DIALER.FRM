VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form DIALER 
   Caption         =   "Marcador telefónico MSComm"
   ClientHeight    =   1545
   ClientLeft      =   4005
   ClientTop       =   3270
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   4800
   WhatsThisHelp   =   -1  'True
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   1920
      TabIndex        =   3
      Top             =   885
      Width           =   975
   End
   Begin VB.CommandButton QuitButton 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3000
      TabIndex        =   1
      Top             =   885
      Width           =   975
   End
   Begin VB.CommandButton DialButton 
      Caption         =   "Marcar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   840
      TabIndex        =   0
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para marcar un número, haga clic en el botón Marcar"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "DIALER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   por Carl Franklin
'
'   Actualizado por Anton de Jong
'
'   Muestra cómo marcar números telefónicos con un módem.
'
'   Para que este programa funcione, el teléfono y el
'   módem deben estar conectados a la misma línea.
'--------------------------------------------------------
Option Explicit

DefInt A-Z

' Este indicador se establece cuando el usuario elige Cancelar.
Dim CancelFlag

Private Sub CancelButton_Click()
    ' CancelFlag indica al procedimiento Dial que termine.
    CancelFlag = True
    CancelButton.Enabled = False
End Sub

Private Sub Dial(Number$)
    Dim DialString$, FromModem$, dummy

    ' AT es el comando compatible Hayes ATTENTION, y es necesario para enviar comandos al módem.
    ' DT significa "marcar con tonos". El comando Dial utiliza tonos, en lugar de pulsos (DP =
    ' marcar con pulsos). Number$ es el número de teléfono a marcar.
    ' El punto y coma indica al módem que vuelva al modo de comando después de marcar
    ' (importante). Es necesario un retorno de carro, vbCr, al enviar comandos al módem.
    DialString$ = "ATDT" + Number$ + ";" + vbCr

    ' Configuración del puerto de comunicaciones.
    ' Suponiendo que hay conectado un mouse a COM1, se establece 2 en CommPort
    MSComm1.CommPort = 2
    MSComm1.Settings = "9600,N,8,1"
    
    ' Abre el puerto de comunicaciones.
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2 no está disponible. Cambie la propiedad CommPort a otro puerto."
       Exit Sub
    End If
    
    ' Vuelca el búfer de entrada.
    MSComm1.InBufferCount = 0
    
    ' Marca el número.
    MSComm1.Output = DialString$
    
    ' Espera a que venga el "OK" del módem.
    Do
       dummy = DoEvents()
       ' Si hay datos en el búfer, los lee.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Comprueba si hay "OK".
          If InStr(FromModem$, "Aceptar") Then
             ' Indica al usuario que descuelgue el teléfono.
             Beep
             MsgBox "Escriba el número y presione Intro o haga clic en Aceptar"
             Exit Do
          End If
       End If
        
       ' ¿El usuario ha elegido Cancelar?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Desconecta el módem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Cierra el puerto.
    MSComm1.PortOpen = False
End Sub

Private Sub DialButton_Click()
    Dim Number$, Temp$
    
    DialButton.Enabled = False
    QuitButton.Enabled = False
    CancelButton.Enabled = True
    
    ' Obtiene el número a marcar.
    Number$ = InputBox$("Escriba el número de teléfono:", Number$)
        If Number$ = "" Then Exit Sub
    Temp$ = Status
    Status = "Marcando - " + Number$
    
    ' Marca el número de teléfono seleccionado.
    Dial Number$

    DialButton.Enabled = True
    QuitButton.Enabled = True
    CancelButton.Enabled = False

    Status = Temp$
End Sub

Private Sub Form_Load()
    ' Al establecer 0 en InputLen se indica a MSComm que
    ' lea todo el contenido del búfer de entrada cuando
    ' se use la propiedad Input.
    MSComm1.InputLen = 0
    
End Sub

Private Sub QuitButton_Click()
    End
End Sub

