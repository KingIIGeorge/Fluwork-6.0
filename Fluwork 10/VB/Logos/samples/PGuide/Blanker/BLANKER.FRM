VERSION 5.00
Begin VB.Form DemoForm 
   BackColor       =   &H00000000&
   Caption         =   "Demostración de protector de pantalla"
   ClientHeight    =   4425
   ClientLeft      =   960
   ClientTop       =   1965
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "BLANKER.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4425
   ScaleWidth      =   7470
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   120
   End
   Begin VB.CommandButton cmdStartStop 
      BackColor       =   &H00000000&
      Caption         =   "Comenzar la demostración"
      Default         =   -1  'True
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2430
   End
   Begin VB.PictureBox picBall 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1800
      Picture         =   "BLANKER.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   8
      Left            =   6330
      Picture         =   "BLANKER.frx":0614
      Top             =   3765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line linLineCtl 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   240
      X2              =   4080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   7
      Left            =   5760
      Picture         =   "BLANKER.frx":091E
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   6
      Left            =   5160
      Picture         =   "BLANKER.frx":0C28
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   5
      Left            =   4560
      Picture         =   "BLANKER.frx":0F32
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   4
      Left            =   3960
      Picture         =   "BLANKER.frx":123C
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   3
      Left            =   3360
      Picture         =   "BLANKER.frx":1546
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   2
      Left            =   2760
      Picture         =   "BLANKER.frx":1850
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   1
      Left            =   2160
      Picture         =   "BLANKER.frx":1B5A
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   0
      Left            =   1560
      Picture         =   "BLANKER.frx":1E64
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpClone 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      Height          =   1215
      Index           =   0
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   960
      Top             =   1080
      Width           =   15
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuLineCtlDemo 
         Caption         =   "&Línea saltarina"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCtlMoveDemo 
         Caption         =   "&Rebotes"
      End
      Begin VB.Menu mnuImageDemo 
         Caption         =   "L&una giratoria"
      End
      Begin VB.Menu mnuShapeDemo 
         Caption         =   "&Manicomio"
      End
      Begin VB.Menu mnuPSetDemo 
         Caption         =   "&Confetti"
      End
      Begin VB.Menu mnuLineDemo 
         Caption         =   "&Fuego cruzado"
      End
      Begin VB.Menu mnuCircleDemo 
         Caption         =   "&Tapiz de arco iris"
      End
      Begin VB.Menu mnuScaleDemo 
         Caption         =   "&Barra de colores"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Declara una variable para seguir el marco de la animación.
Dim FrameNum
' Declara las variables de las coordenadas X e Y, que mantienen la posición.
Dim XPos
Dim YPos
' Declara una variable indicadora para los procedimientos de dibujo
' de gráficos en el bucle "Do Loops".
Dim DoFlag
' Declara una variable para seguir los controles de movimiento.
Dim Motion
' Declara las variables de color del formulario.
Dim R
Dim G
Dim B

Private Sub CircleDemo()
    ' Declara variables locales.
    Dim Radius
    ' Crea colores RGB aleatorios.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Sitúa el centro de los círculos en el centro del formulario.
    XPos = ScaleWidth / 2
    YPos = ScaleHeight / 2
    ' Genera un radio comprendido entre 0 y casi la mitad de la altura
    ' del formulario.
    Radius = ((YPos * 0.9) + 1) * Rnd
    ' Dibuja un círculo en el formulario.
    Circle (XPos, YPos), Radius, RGB(R, G, B)
End Sub

Private Sub cmdStartStop_Click()
' Declara variables locales.
Dim UnClone
Dim MakeClone
Dim X1
Dim Y1
    Select Case DoFlag
        Case True
            cmdStartStop.Caption = "Comenzar la demostración"
            DoFlag = False
            mnuOption.Enabled = True
            If mnuCtlMoveDemo.Checked = True Then
                ' Oculta el gráfico otra vez.
                picBall.Visible = False
            ElseIf mnuLineDemo.Checked = True Then
                ' Elimina las líneas del formulario.
                Cls
            ElseIf mnuShapeDemo.Checked = True Then
                ' Elimina todos los controles de tipo "Shape" cargados
                ' dinámicamente.
                For UnClone = 1 To 20
                    Unload shpClone(UnClone)
                Next UnClone
                ' Restablece a negro el color de fondo del formulario.
                DemoForm.BackColor = QBColor(0)
                ' Limpia el formulario para que el cambio de color surta
                ' efecto.
                Refresh
            ElseIf mnuPSetDemo.Checked = True Then
                ' Elimina los trozos de confeti del formulario.
                Cls
            ElseIf mnuLineCtlDemo.Checked = True Then
                ' Oculta otra vez el control de línea.
                linLineCtl.Visible = False
                ' Quita cualquier pixel suelto que haya quedado después de
                ' ocultar la línea.
                Cls
            ElseIf mnuImageDemo.Checked = True Then
                ' Oculta el gráfico saltando otra vez.
                imgMoon(0).Visible = False
            ElseIf mnuScaleDemo.Checked = True Then
                ' Borra el formulario.
                Cls
                ' Recupera la escala por defecto del formulario.
                Scale
            ElseIf mnuCircleDemo.Checked = True Then
                ' Elimina los círculos del formulario.
                Cls
            End If
        Case False
            cmdStartStop.Caption = "Parar la demostración"
            DoFlag = True
            mnuOption.Enabled = False
            If mnuCtlMoveDemo.Checked = True Then
                ' Establece como visible el control cuadro de imagen.
                picBall.Visible = True
                ' Determina el movimiento inicial del gráfico de forma aleatoria.
                ' Los valores predeterminados van de 1 a 4.  El valor de la variable Motion
                ' determina qué parte del bucle "Do Loop" se ejecuta.
                Motion = Int(4 * Rnd + 1)
            ElseIf mnuLineDemo.Checked = True Then
                ' Inicializa el generador de números aleatorios.
                Randomize
                ' Establece el ancho de línea.
                DrawWidth = 2
                ' Establece las coordenadas X e Y a una posición aleatoria dentro del formulario.
                X1 = Int(DemoForm.Width * Rnd + 1)
                Y1 = Int(DemoForm.Height * Rnd + 1)
            ElseIf mnuShapeDemo.Checked = True Then
                ' Carga dinámicamente en el formulario una matriz de 20 controles
                For MakeClone = 1 To 20
                    Load shpClone(MakeClone)
                Next MakeClone
            ElseIf mnuPSetDemo.Checked = True Then
                ' Establece el tamaño de los trozos de confeti.
                DrawWidth = 5
            ElseIf mnuLineCtlDemo.Checked = True Then
                ' Establece como visible el control de línea.
                linLineCtl.Visible = True
                ' Establece el grosor de línea tal como se verá.
                DrawWidth = 7
            ElseIf mnuImageDemo.Checked = True Then
                ' Establece como visible el gráfico de control de imagen.
                imgMoon(0).Visible = True
                ' Establece la imagen inicial de la animación.
                FrameNum = 0
                ' Determina el movimiento inicial del gráfico que salta.
                ' Los valores de configuración van de 1 a 4.  El valor de la variable Motion
                ' determina qué parte del bucle "Do Loop" se ejecuta.
                Motion = Int(4 * Rnd + 1)
            ElseIf mnuScaleDemo.Checked = True Then
                ' Inicializa el generador de números aleatorios.
                Randomize
                ' Establece el contorno del cuadro para que los
                ' cuadros no se superpongan.
                DrawWidth = 1
                ' Establece el valor de la coordenada X al borde izquierdo del formulario.
                ' Establece la coordenada X del primer cuadro a 1, el segundo cuadro a 2,
                ' y así sucesivamente.
                ScaleLeft = 1
                ' Establece la coordenada Y de la parte superior del formulario a 10.
                ScaleTop = 10
                ' Establece el número de unidades de ancho del formulario a un número aleatorio
                ' comprendido entre 3 y 12.  Esto cambia el número de cuadros dibujados cada vez que
                ' comienza el procedimiento.
                ScaleWidth = Int(13 * Rnd + 3)
                ' Establece el número de unidades de altura del formulario a 10.  Así la
                ' altura de todos los cuadros queda comprendida entre 0 y 10, y la
                ' coordenada Y comienza en el borde inferior del formulario.
                ScaleHeight = -10
            ElseIf mnuCircleDemo.Checked = True Then
                ' Define el ancho del contorno del círculo.
                DrawWidth = 1
                ' Dibuja círculos con el estilo de líneas punteadas.
                DrawStyle = vbDash
                ' Dibuja líneas usando el pincel XOR, combinando los colores encontrados en
                ' el pincel o en el monitor, pero no en los dos.
                DrawMode = vbXorPen
            End If
    End Select
End Sub

Private Sub CtlMoveDemo()
    Select Case Motion
    Case 1
        ' Mueve el gráfico a la izquierda y hacia arriba 20 twips mediante el método Move.
        picBall.Move picBall.Left - 20, picBall.Top - 20
        ' Si el gráfico alcanza el borde izquierdo del formulario, se mueve a la
        ' derecha y hacia arriba.
        If picBall.Left <= 0 Then
            Motion = 2
        ' Si el gráfico alcanza el borde superior del formulario, se mueve a la
        ' izquierda y hacia abajo.
        ElseIf picBall.Top <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Mueve el gráfico a la derecha y hacia arriba 20 twips.
        picBall.Move picBall.Left + 20, picBall.Top - 20
        ' Si el gráfico alcanza el borde derecho del formulario, se mueve a la izquierda y
        ' hacia arriba. Este procedimiento determina el borde derecho del formulario restando
        ' el ancho del gráfico del ancho del formulario.
        If picBall.Left >= (DemoForm.Width - picBall.Width) Then
            Motion = 1
        ' Si el gráfico alcanza el borde superior del formulario, se mueve a la derecha y
        ' hacia abajo.
        ElseIf picBall.Top <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Mueve el gráfico a la derecha y hacia abajo 20 twips.
        picBall.Move picBall.Left + 20, picBall.Top + 20
        ' Si el gráfico alcanza el borde derecho del formulario, se mueve a la
        ' izquierda y hacia abajo.
        If picBall.Left >= (DemoForm.Width - picBall.Width) Then
            Motion = 4
        ' Si el gráfico alcanza el borde inferior del formulario, se mueve a la
        ' derecha y hacia arriba. Esta rutina determina el borde inferior del formulario
        ' restando la altura del gráfico de la altura del formulario menos 680 twips
        ' debido a la altura de la barra de título y de la barra de menús.
        ElseIf picBall.Top >= (DemoForm.Height - picBall.Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Mueve el gráfico a la izquierda y hacia abajo 20 twips.
        picBall.Move picBall.Left - 20, picBall.Top + 20
        ' Si el gráfico alcanza el borde izquierdo del formulario, se mueve a la
        ' derecha y hacia abajo.
        If picBall.Left <= 0 Then
            Motion = 3
        ' Si el gráfico alcanza el borde inferior del formulario, se mueve a la
        ' izquierda y hacia arriba.
        ElseIf picBall.Top >= (DemoForm.Height - picBall.Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub Delay()
    Dim Start
    Dim Check
    Start = Timer
    Do Until Check >= Start + 0.15
        Check = Timer
    Loop
End Sub

Private Sub Form_Load()
    DoFlag = False
End Sub

Private Sub Form_Resize()
    If mnuScaleDemo.Checked = True And DemoForm.WindowState = 0 Then
        ' Inicializa el generador de números aleatorios.
        Randomize
        ' Establece el ancho de los contornos del cuadro como estrechos para que los
        ' cuadros no se superpongan.
        DrawWidth = 1
        ' Establece el valor de la coordenada X del lado izquierdo del formulario a 1.
        ' Esto facilita el establecer la posición de cada cuadro.  El primer cuadro
        ' tiene la coordenada X a 1, el segundo tiene la coordenada X a 2, y así
        ' sucesivamente.
        ScaleLeft = 1
        ' Establece el valor de la coordenada Y del borde superior del formulario a 10.
        ScaleTop = 10
        ' Establece el número de unidades del ancho del formulario a un número aleatorio entre
        ' 3 y 12.  Esto cambia el número de cuadros que son dibujados cada vez que el usuario
        ' inicia este procedimiento.
        ScaleWidth = Int(13 * Rnd + 3)
        ' Establece el número de unidades de altura del formulario a -10.  Esto tiene
        ' dos efectos.  El primero, todos los cuadros tendrán una altura que varía de 0 a 10.
        ' El segundo, el valor negativo causa que la coordenada Y empiece en el borde
        ' inferior del formulario en lugar del superior.
        ScaleHeight = -10
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ImageDemo()
    Select Case Motion
    Case 1
        ' Mueve el gráfico a la izquierda y hacia arriba 100 twips mediante el método Move.
        imgMoon(0).Move imgMoon(0).Left - 100, imgMoon(0).Top - 100
        ' Incrementa la animación con la siguiente imagen.
        IncrFrame
        ' Si el gráfico alcanza el borde izquierdo del formulario, se mueve a la derecha
        ' y hacia arriba.
        If imgMoon(0).Left <= 0 Then
            Motion = 2
        ' Si el gráfico alcanza el borde superior del formulario, se mueve a la
        ' izquierda y hacia abajo.
        ElseIf imgMoon(0).Top <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Mueve el gráfico a la derecha y hacia arriba 100 twips.
        imgMoon(0).Move imgMoon(0).Left + 100, imgMoon(0).Top - 100
        ' Incrementa la animación con la siguiente imagen.
        IncrFrame
        ' Si el gráfico alcanza el borde derecho del formulario, se mueve a la izquierda
        ' y hacia arriba. Este procedimiento determina el borde derecho del formulario restando
        ' el ancho del gráfico del ancho del control.
        If imgMoon(0).Left >= (DemoForm.Width - imgMoon(0).Width) Then
            Motion = 1
        ' Si el gráfico alcanza el borde superior del formulario, se mueve a la derecha
        ' y hacia abajo.
        ElseIf imgMoon(0).Top <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Mueve el gráfico a la derecha y hacia abajo 100 twips.
        imgMoon(0).Move imgMoon(0).Left + 100, imgMoon(0).Top + 100
        ' Incrementa la animación con la siguiente imagen.
        IncrFrame
        ' Si el gráfico alcanza el borde derecho del formulario, se mueve a la izquierda
        ' y hacia abajo.
        If imgMoon(0).Left >= (DemoForm.Width - imgMoon(0).Width) Then
            Motion = 4
        ' Si el gráfico alcanza el borde inferior del formulario, se mueve a la derecha y
        ' hacia arriba. Este procedimiento determina el borde inferior del formulario restando
        ' la altura del gráfico de la altura del formulario menos 680 twips debido a la altura
        ' de la barra del título y de la barra de menús.
        ElseIf imgMoon(0).Top >= (DemoForm.Height - imgMoon(0).Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Mueve el gráfico a la izquierda y hacia abajo 100 twips.
        imgMoon(0).Move imgMoon(0).Left - 100, imgMoon(0).Top + 100
        ' Incrementa la animación con la siguiente imagen.
        IncrFrame
        ' Si el gráfico alcanza el borde izquierdo del formulario, se mueve a la derecha y
        ' hacia abajo.
        If imgMoon(0).Left <= 0 Then
            Motion = 3
        ' Si el gráfico alcanza el borde inferior del formulario, se mueve a la izquierda y
        ' hacia arriba.
        ElseIf imgMoon(0).Top >= (DemoForm.Height - imgMoon(0).Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub IncrFrame()
    ' Incrementa el número de imagen.
    FrameNum = FrameNum + 1
    ' La matriz de controles imagen para la animación tiene elementos desde 0 a 7. A
    ' la octava imagen, se restablece el número de imagen a 0 para obtener un bucle de
    ' animación sin fin.
    If FrameNum > 8 Then
        FrameNum = 1
    End If
    ' Establece la propiedad Picture del control de imagen con la propiedad Picture
    ' de la imagen actual.
    imgMoon(0).Picture = imgMoon(FrameNum).Picture
    ' Se genera una pausa para que la animación no sea demasiado rápida.
    Me.Refresh
    Delay
End Sub

Private Sub LineCtlDemo()
    ' Establece las coordenadas X e Y (posición izquierda/derecha) de la posición inicial de
    ' la línea a una posición aleatoria del formulario.
    linLineCtl.X1 = Int(DemoForm.Width * Rnd)
    linLineCtl.Y1 = Int(DemoForm.Height * Rnd)
    ' Establece las coordenadas X e Y (posición izquierda/derecha) de la posición final de la
    ' línea a una posición aleatoria del formulario.
    linLineCtl.X2 = Int(DemoForm.Width * Rnd)
    linLineCtl.Y2 = Int(DemoForm.Height * Rnd)
    ' Borra el formulario para eliminar cualquier pixel residual.
    Cls
    ' Realiza una pausa antes de mover de nuevo la línea.
    Delay
End Sub

Private Sub LineDemo()
    ' Declara variables locales.
    Dim X2
    Dim Y2
    ' Crea colores RGB aleatorios.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Establece el punto de destino del control de línea a una posición aleatoria
    ' en el formulario.
    X2 = Int(DemoForm.Width * Rnd + 1)
    Y2 = Int(DemoForm.Height * Rnd + 1)
    ' Utilizando el método Line, dibuja una línea desde las coordenadas actuales al punto de destino,
    ' dándole un color aleatorio. Cada línea empieza donde acaba la última línea.
    Line -(X2, Y2), RGB(R, G, B)
End Sub

Private Sub mnuCircleDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = True
End Sub

Private Sub mnuCtlMoveDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = True
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuImageDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = True
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuLineCtlDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = True
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuLineDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = True
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuPSetDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = True
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuScaleDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = True
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuShapeDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = True
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub PSetDemo()
    ' Crea colores RGB aleatorios.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' XPos establece la posición horizontal de un bit de confeti a una
    ' posición aleatoria dentro del formulario.
    XPos = Rnd * ScaleWidth
    ' YPos establece la posición vertical de un bit de confeti a una
    ' posición aleatoria dentro del formulario.
    YPos = Rnd * ScaleHeight
    ' Dibuja un bit de confeti en XPos, YPos. Asigna un color aleatorio
    ' al bit de confeti.
    PSet (XPos, YPos), RGB(R, G, B)
End Sub

Private Sub ScaleDemo()
    ' Declara variables locales.
    Dim Box
    ' Crea el mismo número de cuadros que el ancho del formulario.
    For Box = 1 To ScaleWidth
    ' Crea colores RGB aleatorios.
        R = 255 * Rnd
        G = 255 * Rnd
        B = 255 * Rnd
        ' Dibuja cuadros usando el método Line con las opciones B (cuadro) F (relleno).
        ' Los cuadros empiezan en la coordenada X determinada por ScaleWidth y en la
        ' coordenada Y 0 (parte inferior del formulario). Cada cuadro tiene una
        ' anchura de 1 y tiene una altura aleatoria entre 0 y 10. Llena el cuadro con
        ' un color aleatorio.
        Line (Box, 0)-Step(1, (Int(11 * Rnd))), RGB(R, G, B), BF
    Next Box
    ' Realiza una pausa para mostrar todos los cuadros antes de volver a dibujar.
    Delay
End Sub

Private Sub ShapeDemo()
    ' Declara variables locales.
    Dim CloneID
    ' Crea colores RGB aleatorios.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Establece el color de fondo del formulario a un valor aleatorio.
    DemoForm.BackColor = RGB(R, G, B)
    ' Selecciona un control Shape aleatorio de la matriz de controles.
    CloneID = Int(20 * Rnd + 1)
    ' XPos y YPos establecen la posición del control Shape seleccionado
    ' a una posición aleatoria dentro del formulario.
    XPos = Int(DemoForm.Width * Rnd + 1)
    YPos = Int(DemoForm.Height * Rnd + 1)
    ' Establece la figura del control Shape seleccionando una figura aleatoria.
    shpClone(CloneID).Shape = Int(6 * Rnd)
    ' Establece el alto y ancho de un control Shape seleccionado a un tamaño aleatorio entre
    ' 500 y 2500 twips.
    shpClone(CloneID).Height = Int(2501 * Rnd + 500)
    shpClone(CloneID).Width = Int(2501 * Rnd + 500)
    ' Establece el color de fondo y la propiedad DrawMode del control Shape a un color
    ' aleatorio.
    shpClone(CloneID).BackColor = QBColor(Int(15 * Rnd))
    shpClone(CloneID).DrawMode = Int(16 * Rnd + 1)
    ' Mueve el control Shape seleccionado a XPos, YPos.
    shpClone(CloneID).Move XPos, YPos
    ' Establece como visible el control Shape seleccionado.
    shpClone(CloneID).Visible = True
    ' Espera brevemente antes de seleccionar y cambiar el próximo control Shape.
    Delay
End Sub

Private Sub Timer1_Timer()
    If mnuCtlMoveDemo.Checked And DoFlag = True Then
        CtlMoveDemo
    ElseIf mnuLineDemo.Checked And DoFlag = True Then
        LineDemo
    ElseIf mnuShapeDemo.Checked And DoFlag = True Then
        ShapeDemo
    ElseIf mnuPSetDemo.Checked And DoFlag = True Then
        PSetDemo
    ElseIf mnuLineCtlDemo.Checked And DoFlag = True Then
        LineCtlDemo
    ElseIf mnuImageDemo.Checked And DoFlag = True Then
        ImageDemo
    ElseIf mnuScaleDemo.Checked And DoFlag = True Then
        ScaleDemo
    ElseIf mnuCircleDemo.Checked And DoFlag = True Then
        CircleDemo
    End If
End Sub

