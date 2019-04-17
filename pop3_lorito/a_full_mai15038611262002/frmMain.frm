VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POP3 Mail RETR"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   10
      Text            =   "pagotto"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Text            =   "paggoto@192.168.8.125"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "192.168.8.125"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "Invia Messaggio"
      Height          =   615
      Left            =   4815
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadMessage 
      Caption         =   "Leggi Messaggio Selezionato"
      Height          =   615
      Left            =   2895
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCheckMailbox 
      Caption         =   "Controlla Mail"
      Height          =   615
      Left            =   975
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame frame 
      Caption         =   "Messaggi"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6975
      Begin VB.ListBox lstMessages 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin MSWinsockLib.Winsock pop3 
      Left            =   7200
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   4905
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   2595
      TabIndex        =   7
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "POP3 Host:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ArrayBuffer      As Variant

Private Sub cmdCheckMailbox_Click()
    'Controlla le mail
    pop3.Connect txtHost.Text, 110
End Sub

Private Sub cmdDelete_Click()
    'Elimina messaggio
    
End Sub

Private Sub cmdReadMessage_Click()
    'Crea nuovo read message per leggere un messaggio
    
    CreateNewReadMessageForm
End Sub

Private Sub cmdSendMessage_Click()
    'Crea un nuovo form
    
    CreateNewSendMailForm sAdressData
End Sub

Private Sub lstMessages_Click()

End Sub

Private Sub pop3_DataArrival(ByVal bytesTotal As Long)
    

    Dim strData As String
    Static intMessages          As Integer 'Il n di messaggi da caricare
    
    Static intCurrentMessage    As Integer 'Il contatore di messaggi
    
    Static strBuffer            As String  'Il buffer
    
    'Salvo i dati in una stringa
    
    pop3.GetData strData
    Debug.Print strData
    If Left$(strData, 1) = "+" Or pop3state = POP3_RETR Then
        'Se il primo carattere della risposta è "+" allora
        'Il server accetta la richiesta
        Select Case pop3state
            Case POP3_Connect
                'Resetta il numero di messaggi ricevuti
                intMessages = 0
                pop3state = POP3_USER
                'Invio l username
                pop3.SendData "USER " & txtUsername.Text & vbCrLf
                Debug.Print "USER " & txtUsername.Text
            Case POP3_USER
                'Invio la passowrd
                pop3state = POP3_PASS
                pop3.SendData "PASS " & txtPass.Text & vbCrLf
                Debug.Print "PASS " & txtPass.Text
            Case POP3_PASS
                pop3state = POP3_STAT
                'Invio il comand STAT per sapere quanti n messaggi ci sono
                pop3.SendData "STAT" & vbCrLf
                Debug.Print "STAT"
            Case POP3_STAT
            
                Me.Caption = "STAT CALLED"
                
                'La risposta a STAT potrebbe essere :
                '"+OK 0 0" (no messages at the mailbox) oppure "+OK 3 7564"
                intMessages = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                If intMessages > 0 Then
                    'Le mail nella mail box
                    pop3state = POP3_RETR
                    'Incremento il numero di mail count
                    intCurrentMessage = intCurrentMessage + 1
                    'invio RETR ricevi comando al server
                    'gli dico di ricevere il primo messaggio
                    pop3.SendData "RETR 1" & vbCrLf '1 per il primo messaggio
                    Me.Caption = "RETR 1"
                Else
                    'La mail box è vuota.
                    pop3state = POP3_QUIT
                    pop3.SendData "QUIT" & vbCrLf
                    Me.Caption = "QUIT"
                    Me.Caption = "Non hai email."
                End If
            Case POP3_RETR
                'In questa parte si ricevono i messaggi col RETR
                strBuffer = strBuffer & strData
                'Se in una stringa c'è un punto (.) allora è un messaggio
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    'Elimina la prima parte di stringa del server
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    'Aggiungo il messaggio alla lstMessages
                    ArrayBuffer = SplitMessage(strBuffer)
                    MessageBuffer(intCurrentMessage - 1) = ArrayBuffer(3)
                    FromBuffer(intCurrentMessage - 1) = Trim(ArrayBuffer(2))
                    SubjectBuffer(intCurrentMessage - 1) = ArrayBuffer(0)
                    lstMessages.AddItem "oggetto" 'SubjectBuffer(intCurrentMessage - 1)
                    'Svuoto il buffer per nuovi messaggi
                    strBuffer = ""
                    If intCurrentMessage = intMessages Then
                        'Se è l ultimo messaggio allora chiudi la connessione
                        pop3state = POP3_QUIT
                        pop3.SendData "QUIT" & vbCrLf
                        Debug.Print "QUIT"
                    Else
                        intCurrentMessage = intCurrentMessage + 1
                        pop3state = POP3_RETR
                        'Uso RETR NUMERO per ricevere le mail una alla volta tutte quante
                        pop3.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        Debug.Print "RETR " & intCurrentMessage
                    End If
                End If
            Case POP3_QUIT
                'chiudo la connessione
                pop3.Close
                'gestisco i messaggi
                
        End Select
    Else
        'Se ce un errore
            pop3.Close
            Debug.Print "POP3 Error: " & strData
    End If
End Sub

Private Sub pop3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Ce stato un errore
    MsgBox "Error: #" & Number & vbCrLf & Description
End Sub
