VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalculadoraExcel 
   Caption         =   "Calculadora no Excel"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "frmCalculadoraExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalculadoraExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim primeiroNum As Double
Dim segundoNum As Double
Dim resultado As Double
Dim operador As String

Private Sub LabelAdicao_Click()

    operador = "+"
    
    On Error GoTo trataErro
    
    primeiroNum = TextBoxDisplay.Text
    
    If resultado > 0 Then
        LabelEstrutura.Caption = TextBoxDisplay.Text + operador
    Else
        LabelEstrutura.Caption = LabelEstrutura + TextBoxDisplay.Text + operador
    End If
    
    TextBoxDisplay.Text = ""
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "mais", speakasync:=True
    End If
    
    Exit Sub

trataErro:
    LabelEstrutura.Caption = ""
    primeiroNum = 0
    
End Sub

Private Sub LabelSubtracao_Click()

    operador = "-"
    
    On Error GoTo trataErro
    
    primeiroNum = TextBoxDisplay.Text
    
    If resultado > 0 Then
        LabelEstrutura.Caption = TextBoxDisplay.Text + operador
    Else
        LabelEstrutura.Caption = LabelEstrutura + TextBoxDisplay.Text + operador
    End If
    
    TextBoxDisplay.Text = ""
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "menos", speakasync:=True
    End If
    
    Exit Sub
    
trataErro:
    LabelEstrutura.Caption = ""
    primeiroNum = 0
    
End Sub

Private Sub LabelMultiplicacao_Click()

    operador = "*"
    
    On Error GoTo trataErro
    
    primeiroNum = TextBoxDisplay.Text
    
    If resultado > 0 Then
        LabelEstrutura.Caption = TextBoxDisplay.Text + operador
    Else
        LabelEstrutura.Caption = LabelEstrutura + TextBoxDisplay.Text + operador
    End If
    
    TextBoxDisplay.Text = ""
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "vezes", speakasync:=True
    End If
    
    Exit Sub
    
trataErro:
    LabelEstrutura.Caption = ""
    primeiroNum = 0
    
End Sub

Private Sub LabelDivisao_Click()

    operador = "/"
    
    On Error GoTo trataErro
    
    primeiroNum = TextBoxDisplay.Text
    
    If resultado > 0 Then
        LabelEstrutura.Caption = TextBoxDisplay.Text + operador
    Else
        LabelEstrutura.Caption = LabelEstrutura + TextBoxDisplay.Text + operador
    End If
    
    TextBoxDisplay.Text = ""
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "dividido por", speakasync:=True
    End If
    
    Exit Sub
    
trataErro:
    LabelEstrutura.Caption = ""
    primeiroNum = 0
    
End Sub

Private Sub LabelIgual_Click()
    
    On Error GoTo trataErro
    
    segundoNum = TextBoxDisplay.Text
    
    If operador = "+" Then
        resultado = primeiroNum + segundoNum
    ElseIf operador = "-" Then
        resultado = primeiroNum - segundoNum
    ElseIf operador = "*" Then
        resultado = primeiroNum * segundoNum
    ElseIf operador = "/" Then
        resultado = primeiroNum / segundoNum
    End If
    
    LabelEstrutura.Caption = LabelEstrutura.Caption + TextBoxDisplay.Text
    
    TextBoxDisplay.Text = resultado
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "igual a" & resultado, speakasync:=True
    End If
    
    Exit Sub
    
trataErro:
    segundoNum = 0
    
End Sub

Private Sub LabelZero_Click()
    
    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "0"
    Else
        TextBoxDisplay.Text = TextBoxDisplay.Text + "0"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "zero", speakasync:=True
    End If
    
End Sub

Private Sub LabelUm_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "1"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "1"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "um", speakasync:=True
    End If
    
End Sub

Private Sub LabelDois_Click()
    
    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "2"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "2"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "dois", speakasync:=True
    End If
    
End Sub

Private Sub LabelTres_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "3"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "3"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "três", speakasync:=True
    End If
    
End Sub

Private Sub LabelQuatro_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "4"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "4"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "quatro", speakasync:=True
    End If

End Sub

Private Sub LabelCinco_Click()

   If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "5"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "5"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "cinco", speakasync:=True
    End If

End Sub

Private Sub LabelSeis_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "6"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "6"
    End If

    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "seis", speakasync:=True
    End If
End Sub

Private Sub LabelSete_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "7"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "7"
    End If

    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "sete", speakasync:=True
    End If
End Sub

Private Sub LabelOito_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "8"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "8"
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "oito", speakasync:=True
    End If
End Sub

Private Sub LabelNove_Click()

    If TextBoxDisplay.Text = "0" Then
        TextBoxDisplay.Text = "9"
    Else
        TextBoxDisplay.Text = TextBoxDisplay + "9"
    End If

'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "nove", speakasync:=True
    End If
End Sub

Private Sub LabelC_Click()

    TextBoxDisplay.Text = 0
    LabelEstrutura.Caption = ""
    primeiroNum = 0
    segundoNum = 0
    resultado = 0
    operador = 0
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "limpar", speakasync:=True
    End If
    
End Sub

Private Sub LabelDelete_Click()

    If TextBoxDisplay <> Empty Then
        TextBoxDisplay.Text = Left(TextBoxDisplay.Text, Len(TextBoxDisplay.Text) - 1)
    End If
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "apagar", speakasync:=True
    End If
End Sub

Private Sub LabelPonto_Click()

    If InStr(1, TextBoxDisplay.Text, ",") = 0 Then
        TextBoxDisplay.Text = TextBoxDisplay.Text + ","
    End If

'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "vírgula", speakasync:=True
    End If
End Sub

Private Sub LabelPorcentagem_Click()

    Dim formula1 As Double
    Dim formula2 As Double
    
    segundoNum = TextBoxDisplay.Text
    
    formula1 = primeiroNum * segundoNum / 100 'Adição e Subtração
    formula2 = segundoNum / 100 'Multiplicação e Divisão
    
    If operador = "+" Then
        resultado = primeiroNum + formula1
        LabelEstrutura.Caption = LabelEstrutura.Caption & formula1
    ElseIf operador = "-" Then
        resultado = primeiroNum - formula1
        LabelEstrutura.Caption = LabelEstrutura.Caption & formula1
    ElseIf operador = "*" Then
        resultado = primeiroNum * formula2
        LabelEstrutura.Caption = LabelEstrutura.Caption & formula2
    ElseIf operador = "/" Then
        resultado = primeiroNum / formula2
        LabelEstrutura.Caption = LabelEstrutura.Caption & formula1
    End If
    
    TextBoxDisplay.Text = resultado
    
    'Sound
    If CheckBoxSound.Value = True Then
        Application.Speech.Speak "porcento, igaual a" & resultado, speakasync:=True
    End If
    
End Sub

'Sound
Private Sub CheckBoxSound_Click()

    If CheckBoxSound.Value = True Then
        CheckBoxSound.Caption = "Sound On"
    Else
        CheckBoxSound.Caption = "Sound Off"
     End If

End Sub

Private Sub UserForm_Initialize()
    
    'Som desativado
    CheckBoxSound.Value = False
    CheckBoxSound.Caption = "Sound Off"

End Sub
