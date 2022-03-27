[3/18 10:22 AM] THATIANE ROSA GOMES DEBOLETO
Option ExplicitSub exemplo1()
Dim nome As String
nome = "thatiane"
MsgBox "o nome é: " & nome
End Sub
Sub exemplo2()
Dim n1 As Integer
Dim n2 As Integer
Dim resultado As Double
n1 = 11
n2 = 2
resultado = n1 / n2
MsgBox "resultado é: " & resultadoEnd SubSub exemplo3()
Dim salario As Currency
salario = 5560.89
Dim aumento As Double
aumento = 0.11
Dim salarioNovo As Currency
salarioNovo = salario + (salario * aumento)
MsgBox " o novo salario é:" & Format(salarioNovo, "currency")
End Sub
Sub exemplo4()
Dim dataNasc As Date
Dim hora As Date
dataNasc = #12/3/2001#
hora = Now
MsgBox "data de nascimento da pessoa:" & dataNasc
MsgBox "hora atual: " & Format(hora, " hh:mm:ss am/pm")
End Sub
Sub exemplo5()
Dim valor As Long
Dim valor2 As Long
Dim soma As Long
'para entrada precisamos de um input
valor = InputBox("digite um valor")
valor2 = InputBox("digite outro valor")
soma = valor + valor2
MsgBox "resultado da soma é: " & somaEnd SubSub exemplo6()Dim mensagem As String
Dim nome As String
Dim dataNasc As Date
Dim email As String
Dim salario As Currency
nome = InputBox("digite o nome completo")
dataNasc = InputBox("digite dd/mm/aaaa do nascimento")
email = InputBox("Digite o email")
salario = InputBox("Digite o salario ")
mensagem = mensagem & Chr(10) & Chr(10)
mensagem = mensagem & "Nome: " & nome & Chr(10)
mensagem = mensagem & "Data Nascimento: " & dataNasc & Chr(10)
mensagem = mensagem & "Email: " & email & Chr(10)
mensagem = mensagem & "Salário: " & Format(salario, "currency")
MsgBox mensagem, vbInformation, "Ficha Completa"
End Sub

[3/18 10:58 AM] THATIANE ROSA GOMES DEBOLETO
Sub exemplo7()
Dim filhos As Integer
filhos = MsgBox(" você possui filhos?", vbYesNo, "filhos")
If filhos = vbYes Then
MsgBox " beleza, voce tem filhos", vbInformation, "sucesso"
Else
MsgBox " que pena você não tem filhos", vbExclamation, "tente ter"
End IfEnd Sub
Sub exemplo8()
Static contador As Integer
contador = contador + 1
MsgBox "vez:" & contadorEnd SubSub exemplo9()
Dim valor As Double
valor = 2 ^ 4 - 7 * 2 / 2
MsgBox valor
valor = 2 ^ (4 - (7 * 2) / 2)
MsgBox valorEnd SubSub exemplo10()
Dim casado As Boolean
casado = True
If Not casado Then
MsgBox "não esta casado"
Else
MsgBox "parabens voce tem varias contas a pagar"
End IfEnd Sub

