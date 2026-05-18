VERSION 5.00
Begin VB.UserControl ShapeLabel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ClipControls    =   0   'False
   ForwardFocus    =   -1  'True
   PropertyPages   =   "CPShapeL.ctx":0000
   ScaleHeight     =   1575
   ScaleWidth      =   2820
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   555
   End
   Begin VB.Shape shpBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "ShapeLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
' Notas sobre ShapeLabel:
'
' 1) Casi todo este código se ha generado con
'    el Asistente de interfaz de controles
'    ActiveX. Los puntos donde se ha agregado
'    o modificado manualmente el código se indican
'    con comentarios.
' 2) En la propiedad ForwardFocus del UserControl
'    se ha establecido True para que el enfoque
'    se comporte como en los controles Label
'    ordinarios.
'

Const RESIZE_AdjustX As Single = 0.07
Const RESIZE_AdjustY As Single = 0.03

' Almacenamiento para valores de propiedades. (La
'   mayoría de los valores de las propiedades de
'   ShapeLabel se almacenan en las propiedades del
'   UserControl o de sus controles constituyentes).
Private m_Alignment As AlignmentConstants

' Declaraciones de evento:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

' La propiedad BackColor se ha reasignado
'   manualmente a la propiedad FillColor del
'   control Shape, ya que es el relleno de la
'   forma lo que aparece como fondo de un
'   ShapeLabel.
'
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = shpBack.FillColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    shpBack.FillColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

' ForeColor se asigna a la propiedad ForeColor
'   del control Label, ya que ForeColor de
'   ShapeLabel debe controlar el color de la
'   fuente. El fondo del control Label es
'   Transparent, de modo que BackColor no
'   importa.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

' Para funcionar del mismo modo que las propiedades
'   Enabled de otros controles, Enabled debe tener
'   el ID de procedimiento correcto. El Asistente
'   de interfaz no lo establece, por lo que hay que
'   hacerlo manualmente. Utilice el cuadro de
'   diálogo Atributos del procedimiento, al que se accede
'   desde el menú Herramientas, para establecer
'   Enabled en ID del procedimiento para la
'   propiedad Enabled. El cuadro ID del procedimiento
'   se encuentra en la sección Avanzados del cuadro
'   de diálogo. Seleccione Enabled en el cuadro combinado
'   Nombre para ver los atributos de la propiedad
'   Enabled.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
    ' Agregado manualmente: El cambio de fuente
    '   puede requerir un ajuste en la posición
    '   del control Label.
    Call UserControl_Resize
End Property

' Tipo de la propiedad BorderStyleConstants agregado manualmente.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=shpBack,shpBack,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = shpBack.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    ' shpBack realiza la validación.
    shpBack.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

' Este es un ejemplo de asignación doble.
'   Para que funcione correctamente, el
'   evento Click de ShapeLabel debe
'   producirse cuando el usuario haga clic
'   en la etiqueta, y también cuando lo
'   haga en el cuerpo del control. El
'   Asistente de interfaz no genera código
'   para asignación doble.
Private Sub lblCaption_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

' Asignación para el evento DblClick de
'   Label agregada manualmente.
Private Sub lblCaption_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

' Asignación para el evento MouseDown de
'   Label agregada manualmente.
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Vea el comentario de lblCaption_MouseMove.
    RaiseEvent MouseDown(Button, Shift, _
        ScaleX(X + lblCaption.Left, vbTwips, vbContainerPosition), _
        ScaleY(Y + lblCaption.Height, vbTwips, vbContainerPosition))
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Vea el comentario de UserControl_MouseMove.
    RaiseEvent MouseDown(Button, Shift, _
        ScaleX(X, vbTwips, vbContainerPosition), _
        ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

' Asignación para el evento MouseMove de
'   Label agregada manualmente.
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' La posición del mouse (X, Y) debe pasarse a
    '   las coordenadas del contenedor. Al ser
    '   relativas a Label, primero hay que pasarlas
    '   a coordenadas de UserControl, sumando
    '   lblCaption.Top y .Left. (Pueden sumarse,
    '   pues Scalemode es Twips para UserControl,
    '   al igual que para las coordenadas de Label.
    '   De no ser así, sería necesaria otra
    '   conversión).
    '
    RaiseEvent MouseMove(Button, Shift, _
        ScaleX(X + lblCaption.Left, vbTwips, vbContainerPosition), _
        ScaleY(Y + lblCaption.Height, vbTwips, vbContainerPosition))
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' La posición del mouse (X, Y)debe pasarse a
    '   las coordenadas del contenedor (que
    '   pueden estar en píxeles, por ejemplo).
    '
    ' Observe que podríamos hacer esto de un modo más
    '   general reemplazando vbTwips por
    '   UserControl.Scalemode. Así funcionaría
    '   siempre, incluso cambiando posteriormente
    '   ScaleMode. Sin embargo, ello implicaría dos
    '   llamadas a métodos en lugar de dos constantes.
    '   Suponiendo que MouseMove será tan rápido como
    '   sea posible, se usa vbTwips. (Por supuesto, si
    '   cambia ScaleMode en tiempo de ejecución,
    '   deberá usar UserControl.Scalemode en lugar de
    '   vbTwips).
    '
    RaiseEvent MouseMove(Button, Shift, _
        ScaleX(X, vbTwips, vbContainerPosition), _
        ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

' Asignación para el evento MouseUp de
'   Label agregada manualmente.
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Vea el comentario de lblCaption_MouseMove.
    RaiseEvent MouseUp(Button, Shift, _
        ScaleX(X + lblCaption.Left, vbTwips, vbContainerPosition), _
        ScaleY(Y + lblCaption.Height, vbTwips, vbContainerPosition))
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Vea el comentario de UserControl_MouseMove.
    RaiseEvent MouseUp(Button, Shift, _
        ScaleX(X, vbTwips, vbContainerPosition), _
        ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

' Tipo de la propiedad AlignmentConstants agregado manualmente.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    ' La alineación no se almacena en una
    '   propiedad de control constituyente,
    '   por lo que debemos suministrar
    '   nuestro propio código de validación.
    Select Case New_Alignment
        Case vbCenter
        Case vbLeftJustify
        Case vbRightJustify
        Case Else
            ' Valor de propiedad no válido
            Err.Raise 380
    ' Si se interrumpe aquí la ejecución de
    '   ShapeLabel, haga clic con el botón
    '   derecho en la ventana de código,
    '   seleccione Alternar en el menú de
    '   contexto, y luego seleccione Interrumpción
    '   en errores no controlables. Luego puede
    '   presionar F5 para continuar ejecutando
    '   la demostración.
    End Select
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    ' El cambio de la alineación puede
    '   afectar a las posiciones de los
    '   controles constituyentes.
    Call UserControl_Resize
End Property

' El tipo (OLE_COLOR) de la propiedad
'   BorderColor ha tenido que
'   agregarse manualmente.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=shpBack,shpBack,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = shpBack.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBack.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=shpBack,shpBack,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = shpBack.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    shpBack.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    ' Agregado manualmente: Al cambiar el
    '   título puede ser necesario ajustar
    '   la posición del control Label.
    Call UserControl_Resize
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

' El tipo de la propiedad MousePointerConstants se ha agregado manualmente.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ' UserControl proporciona la validación.
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

' El tipo de la propiedad ShapeConstants se ha agregado manualmente.
'
'¡ADVERTENCIA! ¡NO ELIMINE NI MODIFIQUE LAS LÍNEAS DE COMENTARIO SIGUIENTES!
'MappingInfo=shpBack,shpBack,-1,Shape
Public Property Get Shape() As ShapeConstants
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
    Shape = shpBack.Shape
End Property

Public Property Let Shape(ByVal New_Shape As ShapeConstants)
    ' shpBack proporciona la validación.
    shpBack.Shape() = New_Shape
    PropertyChanged "Shape"
End Property

' Carga los valores de las propiedades desde el almacenamiento.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    shpBack.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    shpBack.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Alignment = PropBag.ReadProperty("Alignment", 2)
    shpBack.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    shpBack.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Caption")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    shpBack.Shape = PropBag.ReadProperty("Shape", 2)
End Sub

Private Sub UserControl_Resize()
    Dim sngCaptionTop As Single
    Dim sngCaptionLeft As Single
    
    ' El control Shape que proporciona el
    '   fondo de ShapeLabel cambia su tamaño
    '   para cubrir todo el control.
    shpBack.Move 0, 0, ScaleWidth, ScaleHeight
    ' El control Label que muestra el título de
    '   ShapeLabel se coloca de acuerdo con el
    '   valor de la propiedad Alignment.
    Select Case Alignment
        Case vbCenter
            sngCaptionLeft = (ScaleWidth - lblCaption.Width) / 2
        Case vbLeftJustify
            sngCaptionLeft = RESIZE_AdjustX * ScaleWidth
        Case vbRightJustify
            sngCaptionLeft = ScaleWidth - lblCaption.Width - RESIZE_AdjustX * ScaleWidth
    End Select
    ' Una propiedad VerticalAlignment
    '   funcionaría de modo similar, y
    '   requeriría su propio Enum.
    sngCaptionTop = (ScaleHeight - lblCaption.Height) / 2 - RESIZE_AdjustY * ScaleHeight
    '
    lblCaption.Move sngCaptionLeft, sngCaptionTop
End Sub

' Escribe los valores de las propiedades en el almacenamiento.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", shpBack.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", shpBack.BorderStyle, 0)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 2)
    Call PropBag.WriteProperty("BorderColor", shpBack.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderWidth", shpBack.BorderWidth, 1)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Shape", shpBack.Shape, 2)
End Sub

