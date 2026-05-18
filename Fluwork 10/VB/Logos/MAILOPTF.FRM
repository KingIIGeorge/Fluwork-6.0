VERSION 5.00
Begin VB.Form MailOptFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de correo"
   ClientHeight    =   3195
   ClientLeft      =   585
   ClientTop       =   870
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MessageOption 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   555
      Left            =   3720
      ScaleHeight     =   555
      ScaleWidth      =   3735
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox retRecip 
         Caption         =   "Confir&mación de lectura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox GeneralOption 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   3915
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3915
      Begin VB.CheckBox DownLoad 
         Caption         =   "&Recuperar correo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3315
      End
      Begin VB.CheckBox NewSess 
         Caption         =   "&Nueva sesión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3195
      End
      Begin VB.CheckBox LogonUI 
         Caption         =   "&Inicio de sesión UI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox UserName 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1980
         TabIndex        =   5
         Top             =   1620
         Width           =   1755
      End
      Begin VB.TextBox MailPassWord 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2040
         Width           =   1755
      End
      Begin VB.CheckBox SendMapiDLG 
         Caption         =   "&Enviar con los diálogos MAPI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   60
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N&ombre de usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1740
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   945
      End
   End
   Begin VB.CommandButton CancelBt 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1980
      TabIndex        =   10
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton OkBt 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   1035
   End
End
Attribute VB_Name = "MailOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBt_Click()
    ' Descargar el formulario actual, y no guardar los cambios.
    Unload Me
End Sub

Private Sub Form_Load()
    ' Configurar los valores iniciales para las casillas de verificación
    ' y los campos de edición.
    UserName = VBMail.MapiSess.UserName
    MailPassWord = VBMail.MapiSess.Password
    NewSession = Abs(VBMail.MapiSess.NewSession)
    LogonUI = Abs(VBMail.MapiSess.LogonUI)
    DownLoadMail = Abs(VBMail.MapiSess.DownLoadMail)
    SendMapiDLG = Abs(SendWithMapi)
    retRecip = Abs(ReturnRequest)
    Select Case OptionType
        Case conOptionMessage
            Call SetupOptionForm(MessageOption)
        Case conOptionGeneral
            Call SetupOptionForm(GeneralOption)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case OptionType
        Case conOptionMessage
            Call SetupOptionForm(MessageOption)
        Case conOptionGeneral
            Call SetupOptionForm(GeneralOption)
    End Select
End Sub

Private Sub OkBt_Click()
    ' Guardar los cambios del usuario.
    ' Característica adicional: Almacena los valores en un archivo .INI.
    VBMail.MapiSess.UserName = UserName
    VBMail.MapiSess.Password = MailPassWord
    VBMail.MapiSess.NewSession = NewSession
    VBMail.MapiSess.LogonUI = LogonUI
    VBMail.MapiSess.DownLoadMail = DownLoadMail
    SendWithMapi = SendMapiDLG
    ReturnRequest = retRecip
    Unload Me
End Sub

Private Sub SetupOptionForm(BasePic As Control)
    BasePic.Top = 0
    BasePic.Left = 0
    BasePic.Visible = True
    BasePic.Enabled = True
    OkBt.Top = BasePic.Height + 120
    CancelBt.Top = BasePic.Height + 120
    Me.Width = BasePic.Width + 120
    Me.Height = OkBt.Top + OkBt.Height + 495
End Sub

