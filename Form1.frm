VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6075
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "&Clean"
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   4560
      List            =   "Form1.frx":000D
      TabIndex        =   7
      Text            =   "1"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "Form1.frx":001C
      Left            =   4560
      List            =   "Form1.frx":002F
      TabIndex        =   6
      Text            =   "N"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "Form1.frx":0042
      Left            =   4560
      List            =   "Form1.frx":0052
      TabIndex        =   5
      Text            =   "8"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Setup"
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   2160
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0062
      Left            =   4560
      List            =   "Form1.frx":0064
      TabIndex        =   3
      Text            =   "9600"
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0066
      Left            =   4560
      List            =   "Form1.frx":0068
      TabIndex        =   0
      Text            =   "1"
      Top             =   0
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check3_Click()

End Sub

Private Sub Command1_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.Output = Text1.Text + vbCrLf
        Text1.Text = ""
    End If
        
End Sub

Private Sub Command2_Click()
    Dim set_tmp, set_tmp1, set_tmp2, set_tmp3, set_tmp4 As String

    On Error Resume Next

    If (MSComm1.PortOpen = True) Then MSComm1.PortOpen = False

    set_tmp = Combo1.Text
    MSComm1.CommPort = set_tmp

    set_tmp1 = Combo2.Text
    set_tmp2 = Combo3.Text
    set_tmp3 = Combo4.Text
    set_tmp4 = Combo5.Text
    Text1.Text = MSComm1.CommPort

    MSComm1.Settings = set_tmp1 + "," + set_tmp3 + "," + set_tmp2 + "," + set_tmp4
    Text2.Text = MSComm1.Settings + vbCrLf

    MSComm1.PortOpen = True
    
    Select Case Err.Number
    Case 0 '0:正常可開啟
    Text2.Text = "COM " + set_tmp + "可被開啟" + vbCrLf + Text2.Text + vbCrLf
    Case 8005 '8005:Port被佔用
    Text2.Text = "COM " + set_tmp + "目前被佔用" + vbCrLf + Text2.Text + vbCrLf
    Case Else 'Port不存在
    Text2.Text = "COM " + set_tmp + "不存在" + vbCrLf + Text2.Text + vbCrLf
    End Select
    Err.Clear '清除Error Code

End Sub

Private Sub Command3_Click()
    Text2.Text = ""
    
End Sub

Private Sub Form_Load()
    Dim i As Integer

    Combo1.Clear
    Combo2.Clear

    On Error Resume Next
    
    For i = 1 To 16
        MSComm1.CommPort = i
        MSComm1.Settings = "9600,n,8,1"
        MSComm1.PortOpen = True
        
        Select Case Err.Number
        Case 0                          '0:正常可開啟
            Combo1.AddItem i
            Combo1.List(i) = i
            'MSComm1.PortOpen = False
        Case 8005                       '8005:Port被佔用
            Combo1.AddItem i
            Combo1.List(i) = i
            'MSComm1.PortOpen = False
        Case Else                       'Port不存在

        End Select
        MSComm1.PortOpen = False
        Err.Clear                       '清除Error Code
    Next i
    
    
    For i = 0 To 7
    Combo2.AddItem i
    Combo2.List(i) = 300 * (2 ^ i)
    Next i
    
    For i = 0 To 4
    Combo2.AddItem i
    Combo2.List(i + 8) = 57600 * (2 ^ i)
    Next i

    Combo1.ListIndex = 1
    Combo2.ListIndex = 5

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    
End Sub

Private Sub Timer1_Timer()

    Dim value As Integer
    Dim in_str_tmp As String
    
    MSComm1.InputLen = 0

    If MSComm1.InBufferCount Then
        'Rean Data
        in_str_tmp = MSComm1.Input
        
        If in_str_tmp = Chr(13) Then
            Text2.Text = Text2.Text + vbCrLf
        Else
            Text2.Text = Text2.Text + in_str_tmp
        End If
        
    End If
       
End Sub
