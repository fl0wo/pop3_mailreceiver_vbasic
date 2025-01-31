Attribute VB_Name = "modMain"
Public Enum SMTP_State              'Enumerazione per gli stati dell SMTP
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Public Enum POP3States             'Enum per il POP 3
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_QUIT
End Enum

Public Type AdressData             'buffer di risposta
    message         As String
    responseAdress  As String
    subject         As String
End Type

Public sAdressData       As AdressData   'Data address
Public pop3state         As POP3States   'Stato POP3
Public smtpState         As SMTP_State   'Stato SMTP
Public EncodedFile       As String       'Il buffer codificato
Dim frm                  As Form         'Variabile
Public MessageBuffer(50) As String  'Alcuni buffer
Public FromBuffer(50)    As String
Public SubjectBuffer(50) As String

Sub CreateNewReadMessageForm()
    'Crea un fragment per leggere messagi
    Set frm = New frmReadMessage
    frm.Show
End Sub

Sub CreateNewSendMailForm(ByRef Adresses As AdressData)
    'Crea un fragment per inviare mail
    Set frm = New frmSendMail
    frm.Show
    'Riempio la stringa
    frm.txtRecipient = Adresses.responseAdress
    frm.txtMessage = Adresses.message
    frm.txtSubject = "Re: " & Adresses.subject
    'Svuota il buffer
    Adresses.message = ""
    Adresses.responseAdress = ""
    Adresses.subject = ""
End Sub

Public Function SplitMessage(message As String) As Variant
    On Error Resume Next
    Dim Pos             As Long
    Dim Pos2            As Long
    Dim arrx(0 To 3)    As String
    Dim br1             As Long
    Dim br2             As Long
    'Estrai il body del messaggio
    Pos = InStr(1, message, vbCrLf & vbCrLf)
    arrx(3) = Right$(message, Len(message) - Pos - 3)
    'Divido il messaggio in pezzi cosi da gestire meglio
    Splitter = Split(message, vbCrLf)
    'Prendo ogni linea
    For i = 0 To UBound(Splitter)
        'Ogni carattere
        For i2 = 1 To Len(Splitter(i))
            If LCase(Mid(Splitter(i), i2, 8)) = "oggett o:" Then
                'Trovato l oggetto
                'riempio l array
                arrx(0) = Mid(Splitter(i), i2 + 8)
            ElseIf LCase(Mid(Splitter(i), i2, 7)) = "mittente :" Then
                'trovato l indirizzo mittente
                'riempio l array
                arrx(1) = Mid(Splitter(i), 8)
            ElseIf LCase(Mid(Splitter(i), i2, 3)) = "destinatario :" Then
                'trovato l indirizzo destinatario
                'riempio l array
                br1 = InStr(1, Splitter(i), "<")
                br2 = InStr(1, Splitter(i), ">")
                arrx(2) = Mid(Splitter(i), 4)
            End If
        Next i2
    Next i

    If InStr(1, arrx(1), "<") <> 0 Then
        arrx(1) = Replace(arrx(1), "<", " ")
    End If
    If InStr(1, arrx(1), ">") <> 0 Then
        arrx(1) = Replace(arrx(1), ">", " ")
    End If
    If InStr(1, arrx(2), "<") <> 0 Then
        arrx(2) = Replace(arrx(2), "<", " ")
    End If
    If InStr(1, arrx(2), ">") <> 0 Then
        arrx(2) = Replace(arrx(2), ">", " ")
    End If
    'Restituisco l array
    SplitMessage = arrx
End Function
