VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "TERRY M. SITUMORANG - 080821017"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7080
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Choose Server"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   5295
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   5055
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6600
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SEND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LISTEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Last Update:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   69.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         DrawMode        =   15  'Merge Pen Not
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   240
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CURRENT TEMPERATURE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      Height          =   5655
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "2011"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "UNIVERSITAS SUMATERA UTARA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   5295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "FAKULTAS MIPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "FISIKA INSTRUMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer


'Enabled send data
Dim enabledSend As Boolean
Dim enabledWinsock1 As Boolean
Dim enabledWinsock2 As Boolean
Dim winsockUsed As Integer

'URL server
Dim strServer As String
Dim defaultURI As String
Dim eUrl As URL

Dim strData As String

Dim strHTTP As String
Dim X As Integer

Private Sub Combo2_Click()
    Winsock1.Close
    Winsock2.Close
    If Combo2.ListIndex = 0 Then
        strServer = "http://wirasihombing.rumahweb.org/script/update.php"
    ElseIf Combo2.ListIndex = 1 Then
        strServer = "http://testing.hostzi.com/script/update.php"
    ElseIf Combo2.ListIndex = 2 Then
        strServer = "http://terrysitumorang2.tk/script/update.php"
    ElseIf Combo2.ListIndex = 3 Then
        strServer = "http://localhost/tmw/script/update.php"
    End If
    'MsgBox (strServer)
    initialConnection
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "LISTEN" Then
        MSComm1.PortOpen = True
    Else
        MSComm1.PortOpen = False
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "SEND" Then
        enabledSend = True
        Command2.Caption = "STOP"
        Combo2.Enabled = False
    Else
        enabledSend = False
        Command2.Caption = "SEND"
        Combo2.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    Winsock1.Close
    End
End Sub

Private Sub Form_Load()
    i = 30
    enabledSend = False
    strServer = "http://wirasihombing.rumahweb.org/script/update.php"
    'strServer = "http://testing.hostzi.com/script/update.php"
    'strServer = "http://localhost/tmw/script/update.php"
    initialConnection

    enabledWinsock1 = False
    enabledWinsock2 = False
    winsockUsed = 1
    
    Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
    End
End Sub

Private Sub MSComm1_OnComm()
    Dim vkar As String
    Dim status As Byte
    Dim bil As Integer
    
    vkar = MSComm1.Input
    status = Asc(vkar)
    bil = Val(status)
    Label2.Caption = bil
    update
End Sub

Private Sub update()
   If enabledSend Then
        If winsockUsed = 1 Then
            If Not enabledWinsock1 Then
                Winsock1.Connect
                While Not enabledWinsock1
                    DoEvents
                    If Winsock1.State = sckError Then
                        Winsock1.Close
                        Exit Sub
                    End If
                Wend
            End If
            sendData1
            winsockUsed = 2
        Else
            If Not enabledWinsock2 Then
                Winsock2.Connect
                While Not enabledWinsock2
                    DoEvents
                    If Winsock2.State = sckError Then
                        Winsock2.Close
                        Exit Sub
                    End If
                Wend
            End If
            sendData2
            winsockUsed = 1
        End If
    End If
End Sub

Private Sub Winsock1_Close()
    enabledWinsock1 = False
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    enabledWinsock1 = True
    'Text1.Text = Text1.Text & "- CONNECTED "
    Text1.Text = Text1.Text & "#1 Try to sending data at " & Now & " "
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    'Dim strResponse As String
    'Winsock1.GetData strResponse, vbString, bytesTotal
    Text1.Text = Text1.Text & " - SUCCESS at " & Now & vbCrLf
    Label6.Caption = Now
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'MsgBox Description, vbExclamation, "ERROR"
    Text1.Text = Text1.Text & " - FAILED" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub

Sub initialConnection()
    eUrl = ExtractUrl(strServer)
    defaultURI = eUrl.URI
    
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.RemoteHost = eUrl.Host
    
    Winsock2.Protocol = sckTCPProtocol
    Winsock2.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            Winsock1.RemotePort = eUrl.Port
            
            Winsock2.RemotePort = eUrl.Port
        Else
            Winsock1.RemotePort = 80
            
            Winsock2.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        Winsock1.RemotePort = 80
        
        Winsock2.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
End Sub

Sub sendData1()
    strData = URLEncode("temp") & "=" & URLEncode(Label2.Caption)
    
    eUrl.URI = defaultURI & "?" & strData
    
    strHTTP = "GET " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders & vbCrLf
    strHTTP = strHTTP & strPostData & vbCrLf
    strHTTP = strHTTP & "Keep-Alive: 200" & vbCrLf
    strHTTP = strHTTP & "Connection: Keep-Alive" & vbCrLf

    ' send the HTTP request
    'MsgBox (strHTTP
    Winsock1.SendData strHTTP
End Sub

Sub sendData2()
    strData = URLEncode("temp") & "=" & URLEncode(Label2.Caption)
    
    eUrl.URI = defaultURI & "?" & strData
    
    strHTTP = "GET " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders & vbCrLf
    strHTTP = strHTTP & strPostData & vbCrLf
    strHTTP = strHTTP & "Keep-Alive: 200" & vbCrLf
    strHTTP = strHTTP & "Connection: Keep-Alive" & vbCrLf

    ' send the HTTP request
    'MsgBox (strHTTP
    Winsock2.SendData strHTTP
End Sub

Private Sub Winsock2_Close()
    enabledWinsock2 = False
    Winsock2.Close
End Sub

Private Sub Winsock2_Connect()
    enabledWinsock2 = True
    Text1.Text = Text1.Text & "#2 Try to sending data at " & Now & " "
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Label6.Caption = Now
    Text1.Text = Text1.Text & " - SUCCESS at " & Now & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Text1.Text = Text1.Text & " - FAILED" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub
