VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : PUNTOS
'* CONTENIDO     : CLASE PUNTO
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

'****************************************************************************************
' CREAR CLASE VISUAL BASIC
'****************************************************************************************

' AREA DE DATOS
'*******************************************************************
' Declaracion de los Atributos privados para encapsulamiento
Private miNumero As Long
Private miX1 As Long
Private miY1 As Long
Private miColor As Long
Private miTamaño As Long
'Private miPropiedad As String

' Metodos Get y Let para manipulacion de datos de la clase
' Get
Public Property Get Numero() As Long
  Numero = miNumero
End Property
' Let
Public Property Let Numero(ByVal valor As Long)
  miNumero = valor
End Property

' Get
Public Property Get X1() As Long
  X1 = miX1
End Property
' Let
Public Property Let X1(ByVal valor As Long)
  miX1 = valor
End Property

' Get
Public Property Get Y1() As Long
  Y1 = miY1
End Property
' Let
Public Property Let Y1(ByVal valor As Long)
  miY1 = valor
End Property

' Get
Public Property Get Color() As String
  Color = miColor
End Property
' Let
Public Property Let Color(ByVal valor As String)
  miColor = valor
End Property

' Get
Public Property Get Tamaño() As String
  Tamaño = miTamaño
End Property
' Let
Public Property Let Tamaño(ByVal valor As String)
  miTamaño = valor
End Property

' AREA DE METODOS Y FUNCIONALIDAD
'*******************************************************************
' Metodo constructor y destructor de la clase
' Constructor
Private Sub Class_Initialize()
' Que hacer al instanciar el objeto
  miNumero = 0
  miX1 = 0
  miY1 = 0
  miColor = 0
  miTamaño = 5
End Sub
' Destructor
Sub Class_Terminate()
' Que hacer al destruir el objeto
End Sub

' Metodos publicos para la funcionalidad de la clase
Public Sub Mover(ByVal despla As Long, ByVal direc As Long)
  Select Case direc
  Case 1
    ' mostrar arriba
    If Y1 > 200 Then
      Y1 = Y1 - despla
    End If
  Case 2
    ' mostrar abajo
    If Y1 < 7400 Then
      Y1 = Y1 + despla
    End If
  Case 3
    ' mostrar derecha
    If X1 < 7400 Then
      X1 = X1 + despla
    End If
  Case 4
    ' mostrar izquierda
    If X1 > 200 Then
      X1 = X1 - despla
    End If
  Case Else
  End Select
End Sub

Public Function miFuncion() As String
' Instrucciones
End Function

' Metodos privados auxiliares para la funcionalidad de la clase
Private Sub miSubrutinaAuxiliar()
' Instrucciones
End Sub
Private Function miFuncionAuxiliar() As String
' Instrucciones
End Function
'*******************************************************************


