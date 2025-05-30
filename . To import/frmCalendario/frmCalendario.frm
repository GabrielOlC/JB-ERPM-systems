VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Calendar"
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4644
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lDiaPrimeiro As Long
Dim lProximo     As Long

Private Sub CommandButton0_Click()
    lsClickBotao CommandButton0.Caption, 0
End Sub

Private Sub CommandButton1_Click()
    lsClickBotao CommandButton1.Caption, 1
End Sub

Private Sub CommandButton10_Click()
    lsClickBotao CommandButton10.Caption, 10
End Sub

Private Sub lsClickBotao(ByVal lValor As String, ByVal lBotao As Long)
    If lBotao <= 6 And lValor > 20 Then
        ActiveCell.Value = CDate(ScrollBar1.Value & "-" & ScrollBar2.Value - 1 & "-" & lValor)
    Else
        If lBotao >= 28 And lValor < 15 Then
            ActiveCell.Value = DateAdd("m", 1, CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-" & lValor))
        Else
            ActiveCell.Value = CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-" & lValor)
        End If
    End If
End Sub

Private Sub CommandButton11_Click()
    lsClickBotao CommandButton11.Caption, 11
End Sub

Private Sub CommandButton12_Click()
    lsClickBotao CommandButton12.Caption, 12
End Sub

Private Sub CommandButton13_Click()
    lsClickBotao CommandButton13.Caption, 13
End Sub

Private Sub CommandButton14_Click()
    lsClickBotao CommandButton14.Caption, 14
End Sub

Private Sub CommandButton15_Click()
    lsClickBotao CommandButton15.Caption, 15
End Sub

Private Sub CommandButton16_Click()
    lsClickBotao CommandButton16.Caption, 16
End Sub

Private Sub CommandButton17_Click()
    lsClickBotao CommandButton17.Caption, 17
End Sub

Private Sub CommandButton18_Click()
    lsClickBotao CommandButton18.Caption, 18
End Sub

Private Sub CommandButton19_Click()
    lsClickBotao CommandButton19.Caption, 19
End Sub

Private Sub CommandButton2_Click()
    lsClickBotao CommandButton2.Caption, 2
End Sub

Private Sub CommandButton20_Click()
    lsClickBotao CommandButton20.Caption, 20
End Sub

Private Sub CommandButton21_Click()
    lsClickBotao CommandButton21.Caption, 21
End Sub

Private Sub CommandButton22_Click()
    lsClickBotao CommandButton22.Caption, 22
End Sub

Private Sub CommandButton23_Click()
    lsClickBotao CommandButton23.Caption, 23
End Sub

Private Sub CommandButton24_Click()
    lsClickBotao CommandButton24.Caption, 24
End Sub

Private Sub CommandButton25_Click()
    lsClickBotao CommandButton25.Caption, 25
End Sub

Private Sub CommandButton26_Click()
    lsClickBotao CommandButton26.Caption, 26
End Sub

Private Sub CommandButton27_Click()
    lsClickBotao CommandButton27.Caption, 27
End Sub

Private Sub CommandButton28_Click()
    lsClickBotao CommandButton28.Caption, 28
End Sub

Private Sub CommandButton29_Click()
    lsClickBotao CommandButton29.Caption, 29
End Sub

Private Sub CommandButton3_Click()
    lsClickBotao CommandButton3.Caption, 3
End Sub

Private Sub CommandButton30_Click()
    lsClickBotao CommandButton30.Caption, 30
End Sub

Private Sub CommandButton31_Click()
    lsClickBotao CommandButton31.Caption, 31
End Sub

Private Sub CommandButton32_Click()
    lsClickBotao CommandButton32.Caption, 32
End Sub

Private Sub CommandButton33_Click()
    lsClickBotao CommandButton33.Caption, 33
End Sub

Private Sub CommandButton34_Click()
    lsClickBotao CommandButton34.Caption, 34
End Sub

Private Sub CommandButton35_Click()
    lsClickBotao CommandButton35.Caption, 35
End Sub

Private Sub CommandButton36_Click()
    lsClickBotao CommandButton36.Caption, 36
End Sub

Private Sub CommandButton37_Click()
    lsClickBotao CommandButton37.Caption, 37
End Sub

Private Sub CommandButton38_Click()
    lsClickBotao CommandButton38.Caption, 38
End Sub

Private Sub CommandButton39_Click()
    lsClickBotao CommandButton39.Caption, 39
End Sub

Private Sub CommandButton4_Click()
    lsClickBotao CommandButton4.Caption, 4
End Sub

Private Sub CommandButton40_Click()
    lsClickBotao CommandButton40.Caption, 40
End Sub

Private Sub CommandButton41_Click()
    lsClickBotao CommandButton41.Caption, 41
End Sub

Private Sub CommandButton42_Click()
    lsClickBotao CommandButton42.Caption, 42
End Sub

Private Sub CommandButton43_Click()
    Unload Me
End Sub

Private Sub OkButton_Click()
    Unload Me
End Sub

Private Sub CommandButton5_Click()
    lsClickBotao CommandButton5.Caption, 5
End Sub

Private Sub CommandButton6_Click()
    lsClickBotao CommandButton6.Caption, 6
End Sub

Private Sub CommandButton8_Click()
    lsClickBotao CommandButton8.Caption, 8
End Sub

Private Sub CommandButton9_Click()
    lsClickBotao CommandButton9.Caption, 9
End Sub

Private Sub lblHoje_Click()
    lsClickBotao Day(lblHoje.Caption), 0
End Sub

Private Sub ScrollBar1_Change()
    lsFontesBotoes Me
    txtAno.Text = CStr(ScrollBar1.Value)
    lsAbertura
End Sub

Private Sub ScrollBar1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub ScrollBar2_Change()
    lsFontesBotoes Me
    lblMes = Format(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01", "mmmm")
    lsAbertura
End Sub

Public Sub lsAbertura()
    lDiaPrimeiro = Format(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01"), "w")
    
    Select Case lDiaPrimeiro
        Case 1
            CommandButton0.Caption = 1
            CommandButton1.Caption = 2
            CommandButton2.Caption = 3
            CommandButton3.Caption = 4
            CommandButton4.Caption = 5
            CommandButton5.Caption = 6
            CommandButton6.Caption = 7
            lProximo = 8
        Case 2
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.Caption = 1
            CommandButton2.Caption = 2
            CommandButton3.Caption = 3
            CommandButton4.Caption = 4
            CommandButton5.Caption = 5
            CommandButton6.Caption = 6
            lProximo = 7
        Case 3
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 2)
            CommandButton1.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.ForeColor = &H8000000A
            CommandButton2.Caption = 1
            CommandButton3.Caption = 2
            CommandButton4.Caption = 3
            CommandButton5.Caption = 4
            CommandButton6.Caption = 5
            lProximo = 6
        Case 4
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 3)
            CommandButton1.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 2)
            CommandButton2.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.ForeColor = &H8000000A
            CommandButton2.ForeColor = &H8000000A
            CommandButton3.Caption = 1
            CommandButton4.Caption = 2
            CommandButton5.Caption = 3
            CommandButton6.Caption = 4
            lProximo = 5
        Case 5
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 4)
            CommandButton1.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 3)
            CommandButton2.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 2)
            CommandButton3.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.ForeColor = &H8000000A
            CommandButton2.ForeColor = &H8000000A
            CommandButton3.ForeColor = &H8000000A
            CommandButton4.Caption = 1
            CommandButton5.Caption = 2
            CommandButton6.Caption = 3
            lProximo = 4
        Case 6
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 5)
            CommandButton1.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 4)
            CommandButton2.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 3)
            CommandButton3.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 2)
            CommandButton4.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.ForeColor = &H8000000A
            CommandButton2.ForeColor = &H8000000A
            CommandButton3.ForeColor = &H8000000A
            CommandButton4.ForeColor = &H8000000A
            CommandButton5.Caption = 1
            CommandButton6.Caption = 2
            lProximo = 3
        Case 7
            CommandButton0.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 6)
            CommandButton1.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 5)
            CommandButton2.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 4)
            CommandButton3.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 3)
            CommandButton4.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 2)
            CommandButton5.Caption = Day(CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01") - 1)
            CommandButton0.ForeColor = &H8000000A
            CommandButton1.ForeColor = &H8000000A
            CommandButton2.ForeColor = &H8000000A
            CommandButton3.ForeColor = &H8000000A
            CommandButton4.ForeColor = &H8000000A
            CommandButton5.ForeColor = &H8000000A
            CommandButton6.Caption = 1
            lProximo = 2
    End Select
    
    lsAtualizarRegistros Me
End Sub


Private Sub txtAno_Change()

End Sub

Private Sub UserForm_Initialize()
    ScrollBar1.Min = Year(Now()) - 50
    ScrollBar1.Value = Year(Now())
    ScrollBar1.Max = Year(Now()) + 50
    
    ScrollBar2.Min = 1
    ScrollBar2.Value = Month(Now())
    ScrollBar2.Max = 12
    
    lblHoje.Caption = Format(Now(), "dd/mm/yyyy")
    
    lsAbertura
End Sub

'Identifica o tipo do objeto e insere se for um dos tipos definidos
Private Sub lsInserir(ByRef lTextBox As Variant, ByVal lValor As Long, ByVal lCor As String)
    If (TypeOf lTextBox Is MSForms.CommandButton) Then
        lTextBox.Caption = lValor
        lTextBox.ForeColor = lCor
    End If
End Sub

Private Sub lsAtualizarRegistros(formulario As UserForm)
    Dim controle    As Control
    Dim lContador   As Long
    Dim lUltimodia  As Long
    Dim lCor        As String
 
    lUltimaLinhaAtiva = Day(DateAdd("m", 1, CDate(ScrollBar1.Value & "-" & ScrollBar2.Value & "-01")) - 1)

    lContador = lProximo
    lCor = vbBlack
 
    For Each controle In formulario.Controls
        If controle.Name = "OkButton" Then
        Else
                If (IsNumeric(Right(controle.Name, 2)) Or Right(controle.Name, 1) >= 7) And TypeOf controle Is MSForms.CommandButton Then
            lsInserir controle, lContador, lCor
            
            If lContador < lUltimaLinhaAtiva Then
                lContador = lContador + 1
            Else
                lContador = 1
                lCor = &H8000000A
            End If
        End If
        End If
    Next
End Sub

Private Sub lsFontesBotoes(formulario As UserForm)
    Dim controle    As Control
    Dim lContador   As Long
    Dim lUltimodia  As Long
  
    For Each controle In formulario.Controls
        If controle.Name = "OkButton" Then
        Else
        If (TypeOf controle Is MSForms.CommandButton) Then
            controle.ForeColor = vbBlack
        End If
        End If
    Next
End Sub
