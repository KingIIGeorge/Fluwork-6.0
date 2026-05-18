VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Top             =   6090
      Width           =   4815
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   3000
      TabIndex        =   22
      Top             =   5610
      Width           =   4815
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   3000
      TabIndex        =   21
      Top             =   5130
      Width           =   4815
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Top             =   4650
      Width           =   4815
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   3000
      TabIndex        =   19
      Top             =   4170
      Width           =   4815
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   3690
      Width           =   4815
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3000
      TabIndex        =   17
      Top             =   3210
      Width           =   4815
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   2730
      Width           =   4815
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Top             =   2250
      Width           =   4815
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   1770
      Width           =   4815
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808080&
      Caption         =   "datos del tecnico"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   40
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   8040
      X2              =   8040
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   8040
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   2880
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   8040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   6600
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form5.frx":0442
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808080&
      Caption         =   $"Form5.frx":0B84
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   840
      TabIndex        =   39
      Top             =   7320
      Width           =   5655
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   2760
      X2              =   2760
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2760
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "tecnico"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If MsgBox("¿Esta seguro de desea Guardar los cambios?", vbYesNo, "Confirmacion") = vbYes Then
Else
Exit Sub
End If

Form5.Command2.Caption = " Aceptar"

Open dbpath + "\datos.dat" For Random As #25 Len = Len(persona)
    
    persona.persona1 = "Seleccionar"
    persona.persona2 = Trim(Text1.Text)
    persona.persona3 = Trim(Text2.Text)
    persona.persona4 = Trim(Text3.Text)
    persona.persona5 = Trim(Text4.Text)
    persona.persona6 = Trim(Text5.Text)
    persona.persona7 = Trim(Text6.Text)
    persona.persona8 = Trim(Text7.Text)
    persona.persona9 = Trim(Text8.Text)
    persona.persona10 = Trim(Text9.Text)
    persona.persona11 = Trim(Text10.Text)
    persona.persona12 = Trim(Text11.Text)
    persona.persona13 = Trim(Text12.Text)
    persona.persona14 = Trim(Text13.Text)
    persona.persona15 = Trim(Text14.Text)
    persona.persona16 = Trim(Text15.Text)
    persona.persona17 = Trim(Text16.Text)
    persona.persona18 = Trim(Text17.Text)
    persona.persona19 = Trim(Text18.Text)
    persona.persona20 = Trim(Text19.Text)
    persona.persona21 = Trim(Text20.Text)
    persona.persona22 = Trim(Text21.Text)
    persona.persona23 = Trim(Text22.Text)
    persona.persona24 = Trim(Text23.Text)
    persona.persona25 = Trim(Text24.Text)
    
Put #25, , persona
Close #25

Form1.Combo1.Clear
Form1.Combo2.Clear

Form1.Combo1.AddItem (persona.persona1)
Form1.Combo2.AddItem (persona.persona1)

If Not Text1.Text = "" Then
    Form1.Combo1.AddItem (Text1.Text)
    Form1.Combo2.AddItem (Text1.Text)
    End If

If Not Text2.Text = "" Then
    Form1.Combo1.AddItem (Text2.Text)
    Form1.Combo2.AddItem (Text2.Text)
    End If

If Not Text3.Text = "" Then
    Form1.Combo1.AddItem (Text3.Text)
    Form1.Combo2.AddItem (Text3.Text)
    End If

If Not Text4.Text = "" Then
    Form1.Combo1.AddItem (Text4.Text)
    Form1.Combo2.AddItem (Text4.Text)
    End If

If Not Text5.Text = "" Then
    Form1.Combo1.AddItem (Text5.Text)
    Form1.Combo2.AddItem (Text5.Text)
    End If

If Not Text6.Text = "" Then
    Form1.Combo1.AddItem (Text6.Text)
    Form1.Combo2.AddItem (Text6.Text)
    End If

If Not Text7.Text = "" Then
    Form1.Combo1.AddItem (Text7.Text)
    Form1.Combo2.AddItem (Text7.Text)
    End If

If Not Text8.Text = "" Then
    Form1.Combo1.AddItem (Text8.Text)
    Form1.Combo2.AddItem (Text8.Text)
    End If

If Not Text9.Text = "" Then
    Form1.Combo1.AddItem (Text9.Text)
    Form1.Combo2.AddItem (Text9.Text)
    End If

If Not Text10.Text = "" Then
    Form1.Combo1.AddItem (Text10.Text)
    Form1.Combo2.AddItem (Text10.Text)
    End If

If Not Text11.Text = "" Then
    Form1.Combo1.AddItem (Text11.Text)
    Form1.Combo2.AddItem (Text11.Text)
    End If

If Not Text12.Text = "" Then
    Form1.Combo1.AddItem (Text12.Text)
    Form1.Combo2.AddItem (Text12.Text)
    End If

persona.persona1 = "Seleccionar"
persona.persona2 = Trim(Text1.Text)
persona.persona3 = Trim(Text2.Text)
persona.persona4 = Trim(Text3.Text)
persona.persona5 = Trim(Text4.Text)
persona.persona6 = Trim(Text5.Text)
persona.persona7 = Trim(Text6.Text)
persona.persona8 = Trim(Text7.Text)
persona.persona9 = Trim(Text8.Text)
persona.persona10 = Trim(Text9.Text)
persona.persona11 = Trim(Text10.Text)
persona.persona12 = Trim(Text11.Text)
persona.persona13 = Trim(Text12.Text)
persona.persona14 = Trim(Text13.Text)
persona.persona15 = Trim(Text14.Text)
persona.persona16 = Trim(Text15.Text)
persona.persona17 = Trim(Text16.Text)
persona.persona18 = Trim(Text17.Text)
persona.persona19 = Trim(Text18.Text)
persona.persona20 = Trim(Text19.Text)
persona.persona21 = Trim(Text20.Text)
persona.persona22 = Trim(Text21.Text)
persona.persona23 = Trim(Text22.Text)
persona.persona24 = Trim(Text23.Text)
persona.persona25 = Trim(Text24.Text)

Form1.Combo1.ListIndex = (0)
Form1.Combo2.ListIndex = (0)

End Sub

Private Sub Command2_Click()

Unload Me

End Sub


Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()

Text1.Text = Trim(persona.persona2)
Text2.Text = Trim(persona.persona3)
Text3.Text = Trim(persona.persona4)
Text4.Text = Trim(persona.persona5)
Text5.Text = Trim(persona.persona6)
Text6.Text = Trim(persona.persona7)
Text7.Text = Trim(persona.persona8)
Text8.Text = Trim(persona.persona9)
Text9.Text = Trim(persona.persona10)
Text10.Text = Trim(persona.persona11)
Text11.Text = Trim(persona.persona12)
Text12.Text = Trim(persona.persona13)
Text13.Text = Trim(persona.persona14)
Text14.Text = Trim(persona.persona15)
Text15.Text = Trim(persona.persona16)
Text16.Text = Trim(persona.persona17)
Text17.Text = Trim(persona.persona18)
Text18.Text = Trim(persona.persona19)
Text19.Text = Trim(persona.persona20)
Text20.Text = Trim(persona.persona21)
Text21.Text = Trim(persona.persona22)
Text22.Text = Trim(persona.persona23)
Text23.Text = Trim(persona.persona24)
Text24.Text = Trim(persona.persona25)

End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = QBColor(7)
Text13.BackColor = QBColor(7)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text1.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = QBColor(15)
Text13.BackColor = QBColor(15)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text10.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text10_LostFocus()
Text10.BackColor = QBColor(15)
Text22.BackColor = QBColor(15)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text11.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text11_LostFocus()
Text11.BackColor = QBColor(15)
Text23.BackColor = QBColor(15)
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text12.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text12_LostFocus()
Text12.BackColor = QBColor(15)
Text24.BackColor = QBColor(15)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text13.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 14 And Len(Text14.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 14 And Len(Text15.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text16.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text17.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text18.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text19.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text20.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text21.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text22.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text23.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text24.Text) > 199 Then KeyAscii = 0
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = QBColor(7)
Text14.BackColor = QBColor(7)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text2.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = QBColor(15)
Text14.BackColor = QBColor(15)
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = QBColor(7)
Text15.BackColor = QBColor(7)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text3.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = QBColor(15)
Text15.BackColor = QBColor(15)
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = QBColor(7)
Text16.BackColor = QBColor(7)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text4.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = QBColor(15)
Text16.BackColor = QBColor(15)
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = QBColor(7)
Text17.BackColor = QBColor(7)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text5.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = QBColor(15)
Text17.BackColor = QBColor(15)
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = QBColor(7)
Text18.BackColor = QBColor(7)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text6.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = QBColor(15)
Text18.BackColor = QBColor(15)
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = QBColor(7)
Text19.BackColor = QBColor(7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text7.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = QBColor(15)
Text19.BackColor = QBColor(15)
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = QBColor(7)
Text20.BackColor = QBColor(7)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text8.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text8_LostFocus()
Text8.BackColor = QBColor(15)
Text20.BackColor = QBColor(15)
End Sub

Private Sub Text9_GotFocus()
Text9.BackColor = QBColor(7)
Text21.BackColor = QBColor(7)
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = QBColor(7)
Text22.BackColor = QBColor(7)
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = QBColor(7)
Text23.BackColor = QBColor(7)
End Sub

Private Sub Text12_GotFocus()
Text12.BackColor = QBColor(7)
Text24.BackColor = QBColor(7)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text9.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Text9_LostFocus()
Text9.BackColor = QBColor(15)
Text21.BackColor = QBColor(15)
End Sub
