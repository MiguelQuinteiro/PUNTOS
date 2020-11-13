VERSION 5.00
Begin VB.Form frmPuntos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "PRIMOS SOBRE CIRCUNFERENCIA"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "Estudio Primos "
      Height          =   7335
      Left            =   7800
      TabIndex        =   20
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtMarca2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   46
         Text            =   "103"
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtMarca1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   44
         Text            =   "53"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtRelacion 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   41
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCuentaSuperior 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtCuentaInferior 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   37
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtDosP 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtModulo 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtRelacionPi 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   29
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtPorcentajePrimos 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   27
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtTotalPrimos 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   25
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtTotalPuntos 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   23
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtPerimetro 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Marca 2"
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Marca 1"
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "     Relación      INF / SUP"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cuenta Superior"
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Cuenta Inferior"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "2 * ( P - 1 )"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "   Primo Medio               - P -"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Relacion con Pi"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje Primos"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Total Primos"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Total Puntos"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Perímetro"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   6720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Controles "
      Height          =   7335
      Left            =   10920
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtLimiteMenos 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   1560
         TabIndex        =   49
         Text            =   "204"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtLimiteMas 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Text            =   "204"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdEjes 
         Caption         =   "Ejes"
         Height          =   495
         Left            =   1560
         TabIndex        =   43
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtCantidadPuntos 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   1560
         TabIndex        =   35
         Text            =   "2"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdMaximo 
         Caption         =   "Máximo"
         Height          =   495
         Left            =   1560
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnimacionMenos 
         Caption         =   "Animacion -"
         Height          =   495
         Left            =   1560
         TabIndex        =   19
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnimacionMas 
         Caption         =   "Animacion +"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCicloMenos 
         Caption         =   "Ciclo -"
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCicloMas 
         Caption         =   "Ciclo + "
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdReinicia 
         Caption         =   "Reinicia"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCirculo 
         Caption         =   "Circulo"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdDesconectaTodo 
         Caption         =   "Desconecta Todo"
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdMostrarTodo 
         Caption         =   "Mostrar Todo"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdConectaTodo 
         Caption         =   "Conecta Todo"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdBorrarTodo 
         Caption         =   "Borrar Todo"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdDesconecta 
         Caption         =   "Desconecta"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdConecta 
         Caption         =   "Conecta"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtPunto 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Text            =   "1"
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAbajo 
         Caption         =   "Abajo"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton cmdIzquierda 
         Caption         =   "Izquierda"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDerecha 
         Caption         =   "Derecha"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdArriba 
         Caption         =   "Arriba"
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdMostar 
         Caption         =   "Mostrar"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : PUNTOS
'* CONTENIDO     : PROGRAMA PRINCIPAL
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

' VARIABLES PUBLICAS
Dim oPunto() As clsPunto
Dim miCantPuntos As Long
Dim miCantPrimos As Long
Dim miPorcentajePrimos As Double
Dim miPi As Double
Dim miRadio As Double
Dim miFactorCircular As Double

Private Sub cmdEjes_Click()
  Line (3800, 0)-(3800, 7600)
  Line (0, 3800)-(7600, 3800)
  Line (0, 0)-(7600, 7600)
  Line (0, 7600)-(7600, 0)
End Sub

Private Sub cmdMaximo_Click()
  txtCantidadPuntos.Text = 650
  Call cmdReinicia_Click
  Call cmdCirculo_Click

  oPunto(1).X1 = 3800
  oPunto(1).Y1 = 3800

  Call cmdMostrarTodo_Click
  Call cmdConecta_Click
End Sub

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()
' Inicialización de variable
  miPi = 3.1415926535
  miRadio = 3750
  miFactorCircular = 0.92

  Call CreaObjetos(Val(txtCantidadPuntos.Text))

  ' Marco
  Line (100, 100)-(7500, 7500), , B

  txtPerimetro.Text = Format((2 * miPi * miRadio), "0.00")
End Sub

' MOSTRAR TODO
Private Sub cmdMostrarTodo_Click()
  Dim p As Long
  ' Coloca al uno en el centro
  oPunto(1).X1 = 3800
  oPunto(1).Y1 = 3800

  For p = 1 To miCantPuntos
    Call Mostrar(oPunto(p))
  Next p
End Sub

' BORRAR TODO
Private Sub cmdBorrarTodo_Click()
  Dim p As Long
  For p = 1 To miCantPuntos
    Call Borrar(oPunto(p))
  Next p
End Sub

' MOSTRAR
Private Sub cmdMostar_Click()
  Call Mostrar(oPunto(Val(txtPunto.Text)))
End Sub

' BORRAR
Private Sub cmdBorrar_Click()
  Call Borrar(oPunto(Val(txtPunto.Text)))
End Sub

' ARRIBA
Private Sub cmdArriba_Click()
'Call Mover(oPunto(Val(txtPunto.Text)), 1)
  Call Borrar(oPunto(Val(txtPunto.Text)))
  oPunto(Val(txtPunto.Text)).Mover 50, 1
  Call Mostrar(oPunto(Val(txtPunto.Text)))
End Sub

' ABAJO
Private Sub cmdAbajo_Click()
'Call Mover(oPunto(Val(txtPunto.Text)), 2)
  Call Borrar(oPunto(Val(txtPunto.Text)))
  oPunto(Val(txtPunto.Text)).Mover 50, 2
  Call Mostrar(oPunto(Val(txtPunto.Text)))
End Sub

' DERECHA
Private Sub cmdDerecha_Click()
'Call Mover(oPunto(Val(txtPunto.Text)), 3)
  Call Borrar(oPunto(Val(txtPunto.Text)))
  oPunto(Val(txtPunto.Text)).Mover 50, 3
  Call Mostrar(oPunto(Val(txtPunto.Text)))
End Sub

' IZQUIERDA
Private Sub cmdIzquierda_Click()
'Call Mover(oPunto(Val(txtPunto.Text)), 4)
  Call Borrar(oPunto(Val(txtPunto.Text)))
  oPunto(Val(txtPunto.Text)).Mover 50, 4
  Call Mostrar(oPunto(Val(txtPunto.Text)))
End Sub

' CONECTA
Private Sub cmdConecta_Click()
  Dim i As Long
  Dim p As Long
  For p = 1 To miCantPuntos
    For i = 1 To miCantPuntos
      If oPunto(i).Color <> 0 And oPunto(Val(txtPunto.Text)).Color <> 0 Then
        'If Primo(i + 2) Then
        Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(Val(txtPunto.Text)).X1, oPunto(Val(txtPunto.Text)).Y1), QBColor(oPunto(i).Color)
        'End If
      Else
        'Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(Val(txtPunto.Text)).X1, oPunto(Val(txtPunto.Text)).Y1), frmPuntos.BackColor
      End If
    Next i
  Next p
End Sub

' DESCONECTA
Private Sub cmdDesconecta_Click()
  Dim i As Long
  Dim p As Long
  For p = 1 To miCantPuntos
    For i = 1 To miCantPuntos
      Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(Val(txtPunto.Text)).X1, oPunto(Val(txtPunto.Text)).Y1), frmPuntos.BackColor
    Next i
  Next p
  Call cmdMostrarTodo_Click
End Sub

' CONECTA TODO
Private Sub cmdConectaTodo_Click()
  Dim i As Long
  Dim p As Long
  For p = 1 To miCantPuntos
    For i = 1 To miCantPuntos
      If oPunto(i).Color <> 0 And oPunto(p).Color <> 0 Then
        Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(p).X1, oPunto(p).Y1), QBColor(oPunto(i).Color)
      Else
        'Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(p).X1, oPunto(p).Y1), frmPuntos.BackColor
      End If
    Next i
  Next p
End Sub

' DESCONECTA TODO
Private Sub cmdDesconectaTodo_Click()
  Dim i As Long
  Dim p As Long
  For p = 1 To miCantPuntos
    For i = 1 To miCantPuntos
      Line (oPunto(i).X1, oPunto(i).Y1)-(oPunto(p).X1, oPunto(p).Y1), frmPuntos.BackColor
    Next i
  Next p
  Call cmdMostrarTodo_Click
End Sub

Private Sub cmdAnimacionMas_Click()
  Open "Relacion.txt" For Output As 1

  txtCantidadPuntos.Text = 3
  Dim i As Long
  For i = 1 To Val(txtLimiteMas.Text)
    '        Dim c As Long
    '        For c = 1 To 100000000
    '        Next c
    Call cmdCicloMas_Click
    DoEvents
  Next i

  Close #1
End Sub

Private Sub cmdAnimacionMenos_Click()
  txtCantidadPuntos.Text = Val(txtLimiteMenos.Text)
  Dim i As Long
  For i = 200 To 1 Step -1
    '        Dim c As Long
    '        For c = 1 To 100000000
    '        Next c
    Call cmdCicloMenos_Click
    DoEvents
  Next i
End Sub

Private Sub cmdCicloMas_Click()
  txtCantidadPuntos.Text = Val(txtCantidadPuntos.Text) + 1
  Call PoneAzul(Val(txtCantidadPuntos.Text))
  'Call cmdReinicia_Click
  Call cmdReinicia_Click
  Call cmdCirculo_Click
  Call cmdMostrarTodo_Click
  Call cmdConecta_Click
  Call CalculaIntervalo
End Sub

Private Sub cmdCicloMenos_Click()
  If Val(txtCantidadPuntos.Text) > 1 Then
    txtCantidadPuntos.Text = Val(txtCantidadPuntos.Text) - 1

    Call PoneAzul(Val(txtCantidadPuntos.Text))

    Call cmdReinicia_Click
    Call cmdCirculo_Click
    Call cmdMostrarTodo_Click
    Call cmdConecta_Click
  End If

End Sub

Private Sub cmdCirculo_Click()
  Dim i As Long
  For i = 1 To miCantPuntos
    oPunto(i).X1 = 3800 + (miRadio * Cos((360 / miCantPuntos) * (miPi / 180) * i) * miFactorCircular)
    oPunto(i).Y1 = 3800 + (miRadio * -Sin((360 / miCantPuntos) * (miPi / 180) * i) * miFactorCircular)
  Next i
End Sub

Private Sub cmdReinicia_Click()
' Borrar todo
  Call cmdDesconectaTodo_Click
  Call cmdBorrarTodo_Click

  ' Destruir objetos
  Dim i As Long
  For i = 1 To miCantPuntos
    oPunto(i).Class_Terminate
  Next i

  ' Construir objetos nuevos
  miCantPuntos = Val(txtCantidadPuntos.Text)
  Call CreaObjetos(miCantPuntos)

  ' Mostrar todo
End Sub

Private Sub Reinicia2()
' Borrar todo
  Call cmdDesconectaTodo_Click
  Call cmdBorrarTodo_Click

  ' Destruir objetos
  Dim i As Long
  For i = 1 To miCantPuntos
    oPunto(i).Class_Terminate
  Next i

  ' Construir objetos nuevos
  miCantPuntos = Val(txtCantidadPuntos.Text)
  Call AgregaObjetos(miCantPuntos)

  ' Mostrar todo
End Sub

' MOSTRAR PUNTO
Public Sub Mostrar(ByRef pP As clsPunto)
'PSet (pP.X1, pP.Y1), QBColor(pP.Color)
  Circle (3800, 3800), 50, vbBlue
  Dim r As Long
  For r = 1 To pP.Tamaño
    'Circle (pP.X1, pP.Y1), pP.Tamaño, QBColor(pP.Color)
    Circle (pP.X1, pP.Y1), r, QBColor(pP.Color)
  Next r
End Sub

' BORRAR PUNTO
Public Sub Borrar(ByRef pP As clsPunto)
'PSet (pP.X1, pP.Y1), frmPuntos.BackColor
  Dim r As Long
  For r = 1 To pP.Tamaño
    Circle (pP.X1, pP.Y1), r, frmPuntos.BackColor
  Next r
End Sub

' MOVER PUNTO
Public Sub Mover(ByRef pP As clsPunto, ByVal pD As Long)
  Dim desplazamiento As Long
  desplazamiento = 50
  Select Case pD
  Case 1
    ' borrar actual
    Call Borrar(pP)
    ' mostrar arriba
    pP.Y1 = pP.Y1 - desplazamiento
    Call Mostrar(pP)
  Case 2
    ' borrar actual
    Call Borrar(pP)
    ' mostrar abajo
    pP.Y1 = pP.Y1 + desplazamiento
    Call Mostrar(pP)
  Case 3
    ' borrar actual
    Call Borrar(pP)
    ' mostrar derecha
    pP.X1 = pP.X1 + desplazamiento
    Call Mostrar(pP)
  Case 4
    ' borrar actual
    Call Borrar(pP)
    ' mostrar izquierda
    pP.X1 = pP.X1 - desplazamiento
    Call Mostrar(pP)
  Case Else
  End Select
End Sub

Public Sub CreaObjetos(ByVal pC As Long)
  miCantPuntos = pC
  miCantPrimos = 0
  ' Cantidad de Objetos
  ReDim oPunto(pC)
  Dim p As Long
  For p = 0 To pC
    ' Creación de Objetos
    Set oPunto(p) = New clsPunto
    With oPunto(p)
      .Numero = p

      .X1 = p * 50
      .Y1 = p * 50
      If Primo(p) = True Then
        .Color = 12
        miCantPrimos = miCantPrimos + 1

        If Val(txtModulo.Text) <> 0 Then
          If p Mod Val(txtModulo.Text) = 0 Then
            .Color = 9
          End If
        End If
      Else
        .Color = 0
      End If
      If p = 1 Then
        .Color = 14
      End If
      If p = Val(txtMarca1.Text) Then
        .Color = 10
      End If
      If p = Val(txtMarca2.Text) Then
        .Color = 10
      End If
    End With
  Next p
  txtTotalPuntos.Text = miCantPuntos
  miPorcentajePrimos = miCantPrimos * 100 / miCantPuntos
  txtTotalPrimos.Text = miCantPrimos
  txtPorcentajePrimos.Text = Format(miPorcentajePrimos, "0.00")
  txtRelacionPi.Text = miPorcentajePrimos * miPi / 50
End Sub

Public Sub AgregaObjetos(ByVal pC As Long)
  miCantPuntos = pC
  'miCantPrimos = 0
  ' Cantidad de Objetos
  ReDim Preserve oPunto(miCantPuntos)
  Dim p As Long
  For p = miCantPuntos To miCantPuntos
    ' Creación de Objetos
    Set oPunto(miCantPuntos) = New clsPunto
    With oPunto(p)
      .Numero = p

      .X1 = p * 50
      .Y1 = p * 50
      If Primo(p) = True Then
        .Color = 12
        miCantPrimos = miCantPrimos + 1
        'txtPunto.Text = p

        If Val(txtModulo.Text) <> 0 Then
          If p Mod Val(txtModulo.Text) = 0 Then
            .Color = 9
          End If
        End If
      Else
        .Color = 0
      End If
      If p = 1 Then
        .Color = 14
      End If
    End With
  Next p
  txtTotalPuntos.Text = miCantPuntos
  miPorcentajePrimos = miCantPrimos * 100 / miCantPuntos
  txtTotalPrimos.Text = miCantPrimos
  txtPorcentajePrimos.Text = Format(miPorcentajePrimos, "0.00")
  txtRelacionPi.Text = miPorcentajePrimos * miPi / 50
End Sub


Private Sub PoneAzul(pA As Long)
  pA = (pA / 2) + 1
  If Primo(pA) Then
    txtModulo.Text = pA
    txtDosP.Text = 2 * (pA - 1)
  End If
End Sub

Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

Public Function CuentaPrimos(ByVal pInf As Long, ByVal pSup As Long) As Long
  Dim i As Long
  CuentaPrimos = 0
  For i = pInf To pSup
    If Primo(i) Then
      CuentaPrimos = CuentaPrimos + 1
    End If
  Next i
End Function

Public Sub CalculaIntervalo()
  txtCuentaInferior.Text = CuentaPrimos(1, Val(txtModulo.Text))
  txtCuentaSuperior.Text = CuentaPrimos(Val(txtModulo.Text), Val(txtDosP.Text))
  txtRelacion.Text = Val(txtCuentaInferior.Text) / Val(txtCuentaSuperior.Text)
  'Print #1, txtRelacion.Text
End Sub


