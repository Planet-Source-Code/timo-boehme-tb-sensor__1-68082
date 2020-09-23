VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   3420
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer5 
      Left            =   240
      Top             =   3480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3480
      Width           =   255
   End
   Begin VB.Timer Timer4 
      Left            =   720
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Left            =   720
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   1095
      Index           =   2
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      Height          =   1095
      Index           =   1
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Index           =   0
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   1095
      Index           =   3
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" ( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As _
    String, _
    ByVal cchBuffer As Long) As Long



Dim OldHeight       As Long
Dim OldWidth        As Long
Dim Merker          As String
Dim MerkerOld       As String
Dim MerkerAnswer    As String
Dim Level           As Integer
Dim MerkerPosition           As Integer


Private Sub Command1_Click()
    Call Start
End Sub

Private Sub Start()
    'Start next level
    Timer5.Enabled = False
    MerkerPosition = 0
    Merker = CalculateMerker(Level, Picture1.UBound)
    MerkerAnswer = ""
    Command1.Enabled = False
    Timer1.Interval = 600 - (Level * 10)
    Timer1.Enabled = True
    If Timer1.Interval - 160 > 0 Then Timer4.Interval = Timer1.Interval - 160
    Label2.Caption = "Level " & Level

End Sub

Private Sub Command2_Click()
    'Show the asked colors
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    'Set Defaults
    Me.Caption = "TB-SENSOR, by Timo Böhme - info@goldengel.ch V1.01"
    Level = 1
    Command1.Caption = "START"
    Command2.Caption = "DEMO"
    Timer2.Interval = "100"
    Timer3.Enabled = False
    Timer3.Interval = "1000"
    Timer5.Enabled = False
    Timer5.Interval = 5000
    Label1.Visible = False
    Label1.Alignment = 2
    Label2.BackStyle = vbTransparent
    Label2.Alignment = 2
    Label2.Caption = "WELCOME TO SENSOR"
    Label3.Caption = "Option append"
    Label3.BackStyle = vbTransparent
    Check1.Caption = ""
    Check1.Value = 1
    Image1.Stretch = True
    Image1.Width = Me.Width
    Image1.Height = Me.Height
    If Dir(App.Path & "\sensor.jpg") <> "" Then Image1.Picture = LoadPicture(App.Path & "\sensor.jpg")
    
    'Because of resizing, we need the AutoRedraw methode
    For i = 0 To Picture1.UBound
        Picture1(i).AutoRedraw = True
    Next
End Sub

Private Sub Form_Resize()
    'Every control checked to resize it
    
    If OldWidth > 0 Then
        Dim Ctl As Control
        Dim T As Object
        For Each Ctl In Me
            If Left(Ctl.Name, 5) <> "Timer" Then
                Ctl.Left = Ctl.Left / (OldWidth / Me.Width)
                Ctl.Top = Ctl.Top / (OldHeight / Me.Height)
                Ctl.Width = Ctl.Width / (OldWidth / Me.Width)
                Ctl.Height = Ctl.Height / (OldHeight / Me.Height)
            End If
        Next
    
    End If
    
    OldWidth = Me.Width
    OldHeight = Me.Height
    
End Sub


Private Function CalculateMerker(ByVal myLevel As Integer, ByVal Max As Integer) As String
    'Creates the colors for the game
    Randomize Timer
    Dim i           As Integer
    Dim W           As String
    
    If Check1.Value = 0 Then
        For i = 0 To myLevel
            W = W & CStr(CInt(Rnd * Max)) & ";"
        Next
    Else
        W = MerkerOld
        W = W & CStr(CInt(Rnd * Max)) & ";"
    End If
    
    CalculateMerker = W
    MerkerOld = W
End Function


Private Sub Picture1_DblClick(Index As Integer)
    'When you click the second time on a Box
    'Windows things you mean a double click.
    If Merker <> "" Then Call Light(Index)

End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Light the box if user is clicking in it
    If Merker <> "" Then Call Light(Index)
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Now check the answer
    If Merker <> "" Then
        Call Light(-1)
        Call CheckAnswer(Index)
    End If
End Sub

Private Sub CheckAnswer(ByVal Index As Integer)
    'Checks if the answer is correct
    MerkerAnswer = MerkerAnswer & Index & ";"
    If MerkerAnswer = Left(Merker, Len(MerkerAnswer)) Then
        Call Answer(True)
        If Len(Merker) = Len(MerkerAnswer) Then
            Call Answer(True)
            Level = Level + 1
            Label2.Caption = "You won this level! congratulation! Next Level=" & Level
            If Level > 100 Then
                MsgBox ("No more levels available")
                Level = Level - 1
            End If
            Command1.Enabled = True
            Merker = ""
            Timer5.Enabled = True
            Command1.SetFocus
        End If
    Else
        Call Answer(False)
        Call MsgBox("Watch again...")
        MerkerAnswer = ""
        Timer1.Enabled = True
    End If
    

End Sub

Private Sub Timer1_Timer()
    'Shows the asked colors of the game
    Dim i As Integer
    
    If MerkerPosition >= UBound(Split(Merker, ";")) Then
        Me.Enabled = True
        Timer1.Enabled = False
        Call Light(-1)
        MerkerPosition = 0
    Else
        Me.Enabled = False
        
        i = Val(Split(Merker, ";")(MerkerPosition))
        Call Light(i)
        MerkerPosition = MerkerPosition + 1
    End If
    
    
End Sub


Private Sub Light(ByVal myIndex As Integer)
    'Let the color light up of the Picturebox
    Dim i       As Integer
    Dim R       As Integer
    Dim G       As Integer
    Dim B       As Integer
    Dim Col1    As Long
    Dim Col2    As Long
    Dim D       As String
    Dim W       As String
    
    
    For i = 0 To Picture1.UBound
        Col1 = Picture1(i).Point(1, 1)
        Call RGBsplit1(Col1, R, G, B)
        If i = myIndex Then
            If Picture1(i).Tag <> "ON" Then
                R = R * 2
                G = G * 2
                B = B * 2
            End If
            Picture1(i).Tag = "ON"
        Else
            If Picture1(i).Tag = "ON" Then
                R = R / 2
                G = G / 2
                B = B / 2
            End If
            Picture1(i).Tag = "OFF"
        End If
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        Picture1(i).BackColor = RGB(R, G, B)
    Next
    
    
    'Play the sound
    W = Space$(260)
    D = App.Path & "\" & "sensor" & myIndex + 1 & ".wav"
    Call GetShortPathName(D, W, Len(W))
    If Len(W) > 0 Then Call PlaySound(W, 0, 1)
    
    'For switching the colors off.
    Timer4.Enabled = False
    Timer4.Enabled = True
End Sub

Private Sub RGBsplit1(ByVal Col As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    'Split the Color in R, G and B value
    B = (Col And 16711680) / 65536
    G = (Col And 65280) / 256
    R = Col And 255
End Sub

Private Sub Timer2_Timer()
    'Intro
    Dim i       As Integer
    
    i = Val(Timer2.Tag)
    If i < 4 Then
        Call Light(i Mod 5)
    ElseIf i < 20 Then
        Call Light(3 - i Mod 4)
    Else
        Call Light(i Mod 4)
    End If
    i = i + 1
    Timer2.Tag = i
    
    If i > 40 Then
        Call Light(-1)
        Timer2.Enabled = False
        Timer2.Tag = ""
    End If
End Sub


Private Sub Answer(ByVal OK As Boolean)
    'Shows the label red or green with text
    If OK Then
        Label1.BackColor = vbGreen
        Label1.Caption = "OK"
    Else
        Label1.BackColor = vbRed
        Label1.Caption = "WRONG"
    End If
    
    Label1.Visible = True
    Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
    'Shows, if the answer was right or wrong
    Label1.Visible = False
    Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
    'Switch off all lights after a while
    Call Light(-1)
End Sub

Private Sub Timer5_Timer()
    Timer5.Enabled = False
    Call Start
End Sub
