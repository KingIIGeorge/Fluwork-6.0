VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Ejemplos con el mouse"
   ClientHeight    =   4395
   ClientLeft      =   240
   ClientTop       =   1785
   ClientWidth     =   7275
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuClickLine 
         Caption         =   "&Clic en una línea"
      End
      Begin VB.Menu mnuScribble 
         Caption         =   "&Garabatos"
      End
      Begin VB.Menu mnuDragDrop 
         Caption         =   "&Arrastrar y colocar"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instrucciones"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuClickLine_Click()
    frmClickLine.Cls
End Sub

Private Sub mnuDragDrop_Click()
    frmDrag.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuInstructions_Click()
    ' Declara variables locales.
    Dim MsgText
    Dim CR
    Dim PB
    ' Inicializa la variable local de ruptura de párrafo.
    CR = Chr$(10) & Chr$(13)
    PB = Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13)
    ' Define el texto del mensaje.
    MsgText = "Clic en una línea"
    MsgText = MsgText & CR & "Haga clic en el formulario varias veces para dibujar líneas conectadas."
    MsgText = MsgText & PB & "Garabatos"
    MsgText = MsgText & CR & "Arrastre el mouse en el formulario para dibujar una línea."
    MsgText = MsgText & PB & "Arrastrar y colocar"
    MsgText = MsgText & CR & "Seleccione un archivo de texto válido(.txt), de mapa de bits (.bmp), ejecutable de Windows (.exe) o Ayuda (.hlp) en el cuadro de lista de archivos. Arrastre el archivo y colóquelo en el cuadro de imagen para visualizar o ejecutar el archivo."
    ' Presenta el mensaje.
    MsgBox MsgText, 64, "Instrucciones de Ejemplos con el mouse"
End Sub

Private Sub mnuScribble_Click()
    frmScribble.Cls
End Sub

