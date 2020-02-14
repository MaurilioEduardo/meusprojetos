

# Enviar e-mails padrão (Com anexos) através do VBA

Public Sub Enviaemail()

Dim emailenvia, emailcorpo, anexo1, anexo2, anexo3, anexo4 As Object
Dim posemail, posnome, posenha, nomempresa, body1, body2, body3, body4, body5, body6 As String

For X = 2 To 600
    Set emailenvia = CreateObject("Outlook.Application")
    Set emailcorpo = emailenvia.CreateItem(0)
    
    posemail = Workbooks("Teste - VBA EMAIL").Sheets("Email, Nome e Senha").Cells(X, 1).Text
    posnome = Workbooks("Teste - VBA EMAIL").Sheets("Email, Nome e Senha").Cells(X, 2).Text
    posenha = Workbooks("Teste - VBA EMAIL").Sheets("Email, Nome e Senha").Cells(X, 3).Text
    nomempresa = Workbooks("Teste - VBA EMAIL").Sheets("Email, Nome e Senha").Cells(2, 4).Text
    
    body1 = "                                                                                          Olá, " & posnome & "!!!" & Chr(13) & Chr(13)
    body2 = nomempresa & " decidiu usar o VExpenses como sua plataforma de prestação de contas." & Chr(13) & Chr(13) & "A intenção é facilitar e agilizar todo o processo!" & Chr(13) & Chr(13)
    body3 = "1- A primeira coisa que você deve fazer é criar uma senha para utilizar o VExpenses, para isso clique no link: " & "https://app.vexpenses.com/login" & " e faça login com as seguintes informações:" & Chr(13)
    body4 = "- Login: " & posemail & Chr(13) & Chr(13)
    body5 = "- Senha: " & posenha & Chr(13) & Chr(13)
    For Z = 1 To 8
        If Z = 1 Then
            body6 = Workbooks("Teste - VBA EMAIL").Sheets("Body").Cells(1, 1).Text & Chr(13)
        Else
            body6 = body6 & Chr(13) & Workbooks("Teste - VBA EMAIL").Sheets("Body").Cells(Z, 1).Text & Chr(13)
        End If
    Next
    
    If posnome = "" Then
        Exit For
    End If
    
    With emailcorpo
    .To = posemail
    .CC = ""
    .BCC = ""
    .Subject = "Nova Plataforma de Reembolsos Corporativos - VExpenses"
    .Body = body1 & body2 & body3 & body4 & body5 & body6
    Set anexo1 = .Attachments.Add("C:/Users/Maurilio/Desktop/[IOS] Guia de uso VExpenses.pdf")
    Set anexo2 = .Attachments.Add("C:/Users/Maurilio/Desktop/Aprovações VExpenses.pdf")
    Set anexo3 = .Attachments.Add("C:/Users/Maurilio/Desktop/[WEB] Guia de uso VExpenses.pdf")
    Set anexo4 = .Attachments.Add("C:/Users/Maurilio/Desktop/[Android] Guia de uso VExpenses.pdf")
    .Send
    End With
    
    newHour = Hour(Now())
    newMinute = Minute(Now()) + 2
    newSecond = Second(Now())
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
    
    Set emailenvia = Nothing
    Set emailcorpo = Nothing
    
Next

End Sub

