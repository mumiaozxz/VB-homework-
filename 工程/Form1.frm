VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13470
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "汽车爬坡度曲线"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   20
      Top             =   8052
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "默认"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      TabIndex        =   19
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "清空"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9600
      TabIndex        =   18
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发动机外特性曲线图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   17
      Top             =   4542
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "功率因数平衡图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   16
      Top             =   7350
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "加速度倒数曲线图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   15
      Top             =   6648
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "汽车功率平衡图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   14
      Top             =   5946
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "驱动力与行驶阻力图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   13
      Top             =   5244
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11640
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11640
      TabIndex        =   10
      Top             =   3030
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11640
      TabIndex        =   6
      Top             =   1170
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11640
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   7815
      Left            =   480
      ScaleHeight     =   7755
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   840
      Width           =   9015
   End
   Begin VB.Label Label7 
      Caption         =   "五档"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "四档"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   3030
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "三档"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "二档"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   5
      Top             =   1170
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "一档"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "主减速器传动比"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9600
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "图像生成"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim I1 As Integer, I2 As Integer, I3 As Integer, I4 As Integer, I5 As Integer


Private Sub Command1_Click()
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0) '坐标轴颜色
Picture1.DrawWidth = 1 '坐标轴线宽
Picture1.Scale (-500, 84)-(5500, -12) '定图幅范围
Picture1.Line (0, 0)-(5000, 0)
Picture1.Line (0, 0)-(0, 75)       '坐标轴画线
Picture1.Line (5000, 0)-(5000, 75)
Picture1.CurrentX = 2400
Picture1.CurrentY = -7
Picture1.Print "n/（r/min)"
Picture1.CurrentX = -150
Picture1.CurrentY = 78
Picture1.Print "Pe/kW"   '标坐标轴含义
Picture1.CurrentX = 4800
Picture1.CurrentY = 78
Picture1.Print "Ttq/(N·m)"
For I = 0 To 5000 Step 500
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 1          '画坐标轴的刻度线
Picture1.Line (I, 1)-(I, 0)
End If
Next
For j = 0 To 5000 Step 500
If j <> 0 Then
Picture1.CurrentX = j - 150
Picture1.CurrentY = -2
Picture1.Print j
Else
Picture1.CurrentX = -100      '标坐标轴的刻度
Picture1.CurrentY = -2
Picture1.Print 0
End If
Next
For K = 0 To 70 Step 10
If K <> 0 Then
Picture1.CurrentX = 60
Picture1.CurrentY = K         '画坐标轴的刻度线
Picture1.Line (60, K)-(0, K)
End If
Next
For K = 0 To 70 Step 10
If K <> 0 Then
Picture1.CurrentX = -200
Picture1.CurrentY = K         '标坐标轴的刻度
Picture1.Print K
End If
Next
For l = 0 To 30 Step 5
If l <> 0 Then
Picture1.CurrentX = 5000
Picture1.CurrentY = l
Picture1.Line (4940, l)-(5000, l) '画坐标轴的刻度线
End If
Next
For l = 0 To 30 Step 5
If l <> 0 Then
Picture1.CurrentX = 5000          '标坐标轴的刻度
Picture1.CurrentY = l
Picture1.Print l * 6
End If
Next

Dim n As Single, Ttq As Single, Pe As Single, T As Single, Pemax As Single, n1 As Single, n2 As Single, Ttqmax As Single

'n=发动机转速，Ttq=发动机转矩，Pe=发动机功率，T=坐标轴刻度，n1，n2=特殊点发动机转速

For n = 600 To 4000 Step 1
Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
T = Ttq / 6

Pe = (Ttq * n) / 9550

Picture1.DrawWidth = 2
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.PSet (n, T)
Picture1.ForeColor = RGB(255, 0, 0)     '画曲线
Picture1.PSet (n, Pe)
Next n
For n = 600 To 4000
Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
If Pe >= Pemax Then
Pemax = Pe
Else
Pemax = Pemax
End If
If Pe = Pemax Then
n1 = n
End If                           '标最值
If Ttq >= Ttqmax Then
Ttqmax = Ttq
Else
Ttqmax = Ttqmax
End If
If Ttq = Ttqmax Then
n2 = n
End If
Next n
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (0, Pemax)-(n1, Pemax)
Picture1.Line (n1, 0)-(n1, Pemax)
Picture1.CurrentX = 3500
Picture1.CurrentY = Pemax + 4
Picture1.Print "Pemax="; Int((Pemax + 0.005) * 100) / 100
Picture1.CurrentX = n1 + 15
Picture1.CurrentY = 5
Picture1.Print "n1="; n1
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (5000, Ttqmax / 6)-(n2, Ttqmax / 6)
Picture1.Line (n2, 0)-(n2, Ttqmax / 6)
Picture1.CurrentX = 2000
Picture1.CurrentY = Ttqmax / 6 + 4
Picture1.Print "Ttqmax="; Int((Ttqmax + 0.005) * 100) / 100
Picture1.CurrentX = n2 + 15
Picture1.CurrentY = 5
Picture1.Print "n2="; n2

End Sub

Private Sub Command2_Click()

Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.DrawWidth = 1
Picture1.Scale (-10, 17)-(135, -3)
Picture1.Line (0, 0)-(125, 0)
Picture1.Line (0, 0)-(0, 15)
Picture1.CurrentX = 60
Picture1.CurrentY = -1.5
Picture1.Print "ua / (km / h)";
Picture1.CurrentX = -2
Picture1.CurrentY = 16
Picture1.Print "F/kN"
For I = 0 To 120 Step 20
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 0.3
Picture1.Line (I, 0.3)-(I, 0)
End If
Next
For j = 0 To 120 Step 20
If j <> 0 Then
Picture1.CurrentX = j - 3
Picture1.CurrentY = -0.3
Picture1.Print j
Else
Picture1.CurrentX = -1
Picture1.CurrentY = -0.5
Picture1.Print 0
End If
Next
For K = 0 To 14 Step 2
If K <> 0 Then
Picture1.CurrentX = 1.3
Picture1.CurrentY = K
Picture1.Line (1.3, K)-(0, K)
End If
Next
For K = 0 To 14 Step 2
If K <> 0 Then
Picture1.CurrentX = -5
Picture1.CurrentY = K
Picture1.Print K
End If
Next

Dim n As Single, Ft As Single, ua As Single, Ttq As Single, Ft1 As Single, Ft2 As Double, I1 As Double, I2 As Double, I3 As Double, I4 As Double, I5 As Double


'n=发动机转速，Ttq=发动机转矩，Ft=汽车驱动力，ua=汽车行驶速度=0.377 *r *n /(Ig * Io),   r=0.367 Io=5.83 Ig=Split(Textg.Text)
'
I1 = Val(Text1)

'For ua = 4 To 22 Step 0.01


For ua = 4 To 18 Step 0.01

n = ua * 5.83 / 0.367 / 0.377 * I1

'4.689?=I1

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I1

'Ft=Ttq *Io *Nt *Ig /r,  r=0.367 Io=5.83 Nt=0.85 Ig=Split(Textg.Text, ",")

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 0) '红
Picture1.PSet (ua, Ft / 1000)
Next ua
'For ua = 6 To 35 Step 0.01
For ua = 6 To 28 Step 0.01

I2 = Val(Text2)

n = ua * 5.83 / 0.367 / 0.377 * I2

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I2

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 255, 0) '黄
Picture1.PSet (ua, Ft / 1000)
Next ua
'For ua = 12 To 57 Step 0.01
For ua = 12 To 40 Step 0.01

I3 = Val(Text3)

n = ua * 5.83 / 0.367 / 0.377 * I3

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I3

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 255) '蓝
Picture1.PSet (ua, Ft / 1000)
Next ua
'For ua = 22 To 84 Step 0.01
For ua = 22 To 70 Step 0.01

I4 = Val(Text4)

n = ua * 5.83 / 0.367 / 0.377 * I4

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I4

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 125, 0) '橙
Picture1.PSet (ua, Ft / 1000)
Next ua
'For ua = 30 To 105 Step 0.01
For ua = 30 To 105 Step 0.01

I5 = Val(Text5)

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I5

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.PSet (ua, Ft / 1000)
Next ua

For ua = 0 To 115 Step 0.01

Ft = 3800 * 9.8 * 0.013 + 2.77 / 21.15 * ua ^ 2

'Ft = Gf + CDA * Ua * Ua / 21.15

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, Ft / 1000)
Next ua

For ua = 30 To 107 Step 0.01

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft1 = Ttq * 5.83 * 0.85 / 0.367 * I5

Ft2 = 3800 * 9.8 * 0.013 + 2.77 / 21.15 * ua ^ 2

If Ft1 >= Ft2 Then
Ft2 = Ft1
uamax = ua
Else
Ft2 = Ft2
uamax = uamax
End If
Next ua
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (uamax, 0)-(uamax, 1.85)
Picture1.CurrentX = uamax + 1
Picture1.CurrentY = 1
Picture1.Print "uamax="; Int((uamax + 0.005) * 100) / 100
Picture1.CurrentX = 11
Picture1.CurrentY = 11.5
Picture1.Print "Ⅰ"
Picture1.CurrentX = 20
Picture1.CurrentY = 7.3
Picture1.Print "Ⅱ"
Picture1.CurrentX = 37
Picture1.CurrentY = 4.5
Picture1.Print "Ⅲ"
Picture1.CurrentX = 45
Picture1.CurrentY = 3.2
Picture1.Print "Ⅳ"
Picture1.CurrentX = 90
Picture1.CurrentY = 2.5
Picture1.Print "Ⅴ"

End Sub


Private Sub Command3_Click()

Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.DrawWidth = 1
Picture1.Scale (-10, 84)-(135, -12)
Picture1.Line (0, 0)-(125, 0)
Picture1.Line (0, 0)-(0, 75)
Picture1.CurrentX = 60
Picture1.CurrentY = -7
Picture1.Print "ua / (km / h)";
Picture1.CurrentX = -2
Picture1.CurrentY = 78
Picture1.Print "Pe/kW"
For I = 0 To 120 Step 20
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 1.5
Picture1.Line (I, 1.5)-(I, 0)
End If
Next
For j = 0 To 120 Step 20
If j <> 0 Then
Picture1.CurrentX = j - 3
Picture1.CurrentY = -2
Picture1.Print j
Else
Picture1.CurrentX = -1
Picture1.CurrentY = -2
Picture1.Print 0
End If
Next
For K = 0 To 70 Step 10
If K <> 0 Then
Picture1.CurrentX = 1.3
Picture1.CurrentY = K
Picture1.Line (1.3, K)-(0, K)
End If
Next
For K = 0 To 70 Step 10
If K <> 0 Then
Picture1.CurrentX = -5
Picture1.CurrentY = K
Picture1.Print K
End If
Next
Dim n As Single, Pe As Single, ua As Single, Ft As Single, I1 As Double, I2 As Double, I3 As Double, I4 As Double, I5 As Double


'n=发动机转速，Pe=发动机功率, Ft=汽车驱动力，ua=汽车行驶速度=0.377 *r *n /(Ig * Io),   r=0.367 Io=5.83 Ig=Split(Textg.Text)

'For ua = 4 To 22 Step 0.001
For ua = 4 To 18 Step 0.01
'For ua = 4 To 20 Step 0.01
I1 = Val(Text1)

n = ua * 5.83 / 0.367 / 0.377 * I1

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Pe = (Ttq * n) / 9550
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 0)
Picture1.PSet (ua, Pe)
Next ua

'For ua = 6 To 38 Step 0.01
For ua = 6 To 28 Step 0.01
'For ua = 6 To 35 Step 0.01
I2 = Val(Text2)

n = ua * 5.83 / 0.367 / 0.377 * I2

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 255, 0)
Picture1.PSet (ua, Pe)
Next ua

'For ua = 10 To 64 Step 0.01
For ua = 10 To 43 Step 0.01
'For ua = 10 To 57 Step 0.01
I3 = Val(Text3)

n = ua * 5.83 / 0.367 / 0.377 * I3

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 255)
Picture1.PSet (ua, Pe)
Next ua

'For ua = 17 To 98 Step 0.01
For ua = 17 To 66 Step 0.01
I4 = Val(Text4)

n = ua * 5.83 / 0.367 / 0.377 * I4

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, Pe)
Next ua

'For ua = 18 To 108 Step 0.01
For ua = 18 To 108 Step 0.01

I5 = Val(Text5)

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 255)
Picture1.PSet (ua, Pe)
Next ua

For ua = 0 To 108 Step 0.01

Ft = 3800 * 9.8 * 0.013 + 2.77 / 21.15 * ua ^ 2

Pe = Ft * ua / 3600 / 0.85

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.PSet (ua, Pe)
Next ua

For ua = 18 To 108 Step 0.01

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe = (Ttq * n) / 9550
If Pe >= Pemax Then
Pemax = Pe
Else
Pemax = Pemax
End If
Next ua
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (0, Pemax)-(110, Pemax)
Picture1.CurrentX = 50
Picture1.CurrentY = Pemax + 3
Picture1.Print "Pemax="; Int((Pemax + 0.005) * 100) / 100

For ua = 30 To 107 Step 0.01

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Pe1 = (Ttq * n) / 9550
Ft = 3800 * 9.8 * 0.013 + 2.77 / 21.15 * ua ^ 2

Pe2 = Ft * ua / 3600 / 0.85

If Pe1 >= Pe2 Then
Pe2 = Pe1
uamax = ua
Else
Pe2 = Pe2
uamax = uamax
End If
Next ua
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (uamax, 0)-(uamax, 62)
Picture1.CurrentX = uamax + 1
Picture1.CurrentY = 4
Picture1.Print "uamax="; Int((uamax + 0.005) * 100) / 100
Picture1.CurrentX = 10
Picture1.CurrentY = 45
Picture1.Print "Ⅰ"
Picture1.CurrentX = 16
Picture1.CurrentY = 45
Picture1.Print "Ⅱ"
Picture1.CurrentX = 25
Picture1.CurrentY = 45
Picture1.Print "Ⅲ"
Picture1.CurrentX = 40
Picture1.CurrentY = 45
Picture1.Print "Ⅳ"
Picture1.CurrentX = 66
Picture1.CurrentY = 48
Picture1.Print "Ⅴ"
Picture1.CurrentX = 72
Picture1.CurrentY = 25
Picture1.Print "(Pf+Pw)/ηT"

End Sub

Private Sub Command4_Click()
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.DrawWidth = 1
Picture1.Scale (-10, 10.5)-(135, -1.5)
Picture1.Line (0, 0)-(125, 0)
Picture1.Line (0, 0)-(0, 9.5)
Picture1.CurrentX = 60
Picture1.CurrentY = -0.8
Picture1.Print "ua / (km / h)";
Picture1.CurrentX = -2
Picture1.CurrentY = 10
Picture1.Print "1/a"
For I = 0 To 120 Step 20
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 0.23
Picture1.Line (I, 0.23)-(I, 0)
End If
Next
For j = 0 To 120 Step 20
If j <> 0 Then
Picture1.CurrentX = j - 3
Picture1.CurrentY = -0.3
Picture1.Print j
Else
Picture1.CurrentX = -1
Picture1.CurrentY = -0.3
Picture1.Print 0
End If
Next
For K = 0 To 9 Step 1
If K <> 0 Then
Picture1.CurrentX = 1.3
Picture1.CurrentY = K
Picture1.Line (1.3, K)-(0, K)
End If
Next
For K = 0 To 9 Step 1
If K <> 0 Then
Picture1.CurrentX = -5
Picture1.CurrentY = K
Picture1.Print K
End If
Next
Dim n As Single, Ttq As Single, ua As Single, Ft As Single, a As Single, I1 As Double, I2 As Double, I3 As Double, I4 As Double, I5 As Double
'n=发动机转速，Ttq=发动机转矩，Ft=汽车驱动力，ua=汽车行驶速度=0.377 *r *n /(Ig * Io),   r=0.367 Io=5.83 Ig=Split(Textg.Text)， a=加速度

'For ua = 3 To 24 Step 0.01
For ua = 3 To 19 Step 0.01

I1 = Val(Text1)

'Text1.Text = 6

n = ua * 5.83 / 0.367 / 0.377 * I1


Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I1

a = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800

'If a >= Amax Then
'Amax = a
'Else
'Amax = Amax
'End If

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 0)
Picture1.PSet (ua, 1 / a)

'Text6.Text = Amax

Next ua

'For ua = 12 To 39 Step 0.01
For ua = 12 To 29 Step 0.01

'Text2.Text = 3.8

I2 = Val(Text2)

n = ua * 5.83 / 0.367 / 0.377 * I2

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12


Ft = Ttq * 5.83 * 0.85 / 0.367 * I2

a = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 255, 0)
Picture1.PSet (ua, 1 / a)
Next ua

'For ua = 20 To 66 Step 0.01
For ua = 20 To 46 Step 0.01

I3 = Val(Text3)

'Text3.Text = 2.5

n = ua * 5.83 / 0.367 / 0.377 * I3

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I3

a = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 255)
Picture1.PSet (ua, 1 / a)
Next ua

'For ua = 29 To 84 Step 0.01
For ua = 30 To 70 Step 0.01

I4 = Val(Text4)

'Text4.Text = 1.6

n = ua * 5.83 / 0.367 / 0.377 * I4
Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I4

a = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.PSet (ua, 1 / a)
Next ua

'For ua = 38 To 92 Step 0.01
For ua = 38 To 95 Step 0.01

I5 = Val(Text5)

'Text5.Text = 1

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I5

a = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 255)
Picture1.PSet (ua, 1 / a)
Next ua
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.CurrentX = 10
Picture1.CurrentY = 0.3
Picture1.Print "Ⅰ"
Picture1.CurrentX = 30
Picture1.CurrentY = 0.6
Picture1.Print "Ⅱ"
Picture1.CurrentX = 50
Picture1.CurrentY = 1.3
Picture1.Print "Ⅲ"
Picture1.CurrentX = 70
Picture1.CurrentY = 3
Picture1.Print "Ⅳ"
Picture1.CurrentX = 90
Picture1.CurrentY = 7
Picture1.Print "Ⅴ"

End Sub

Private Sub Command5_Click()

Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.DrawWidth = 1
Picture1.Scale (-10, 0.4)-(135, -0.05)
Picture1.Line (0, 0)-(125, 0)
Picture1.Line (0, 0)-(0, 0.33)
Picture1.CurrentX = 60
Picture1.CurrentY = -0.03
Picture1.Print "ua / (km / h)";
Picture1.CurrentX = -2
Picture1.CurrentY = 0.36
Picture1.Print "D"
For I = 0 To 120 Step 20
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 0.007
Picture1.Line (I, 0.007)-(I, 0)
End If
Next
For j = 0 To 120 Step 20
If j <> 0 Then
Picture1.CurrentX = j - 3
Picture1.CurrentY = -0.01
Picture1.Print j
Else
Picture1.CurrentX = -1
Picture1.CurrentY = -0.01
Picture1.Print 0
End If
Next
For K = 0 To 0.4 Step 0.1
If K <> 0 Then
Picture1.CurrentX = 1.3
Picture1.CurrentY = K
Picture1.Line (1.3, K)-(0, K)
End If
Next
For K = 0 To 0.4 Step 0.1
If K <> 0 Then
Picture1.CurrentX = -6
Picture1.CurrentY = K + 0.005
Picture1.Print "0"; K
End If
Next

Dim n As Single, Ttq As Single, ua As Single, Ft As Single, D As Single, I1 As Double, I2 As Double, I3 As Double, I4 As Double, I5 As Double

'n=发动机转速，Ttq=发动机转矩，Ft=汽车驱动力，ua=汽车行驶速度=0.377 *r *n /(Ig * Io),   r=0.367 Io=5.83 Ig=Split(Textg.Text)， D=

'For ua = 5 To 22 Step 0.01

For ua = 5 To 18 Step 0.01

I1 = Val(Text1)

n = ua * 5.83 / 0.367 / 0.377 * I1

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I1

D = (Ft - 2.77 / 21.15 * ua * ua) / 3800 / 9.8

Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 0)
Picture1.PSet (ua, D)
Next ua
'For ua = 9 To 37 Step 0.01
For ua = 9 To 28 Step 0.01

I2 = Val(Text2)

n = ua * 5.83 / 0.367 / 0.377 * I2

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I2

D = (Ft - 2.77 / 21.15 * ua * ua) / 3800 / 9.8
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 255, 0)
Picture1.PSet (ua, D)
Next ua

'For ua = 16 To 66 Step 0.01
For ua = 16 To 44 Step 0.01

I3 = Val(Text3)

n = ua * 5.83 / 0.367 / 0.377 * I3

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I3

D = (Ft - 2.77 / 21.15 * ua * ua) / 3800 / 9.8
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 255)
Picture1.PSet (ua, D)
Next ua

'For ua = 22 To 84 Step 0.01
For ua = 22 To 70 Step 0.01

I4 = Val(Text4)

n = ua * 5.83 / 0.367 / 0.377 * I4

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I4
D = (Ft - 2.77 / 21.15 * ua * ua) / 3800 / 9.8
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(255, 0, 255)
Picture1.PSet (ua, D)
Next ua

'For ua = 28 To 108 Step 0.01
For ua = 28 To 102 Step 0.01

I5 = Val(Text5)

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I5

D = (Ft - 2.77 / 21.15 * ua * ua) / 3800 / 9.8
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, D)
Next ua

For ua = 0 To 108 Step 0.01

f = 0.013
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.PSet (ua, f)
Next ua
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.CurrentX = 10
Picture1.CurrentY = 0.31
Picture1.Print "Ⅰ"
Picture1.CurrentX = 20
Picture1.CurrentY = 0.19
Picture1.Print "Ⅱ"
Picture1.CurrentX = 38
Picture1.CurrentY = 0.11
Picture1.Print "Ⅲ"
Picture1.CurrentX = 50
Picture1.CurrentY = 0.08
Picture1.Print "Ⅳ"
Picture1.CurrentX = 90
Picture1.CurrentY = 0.037
Picture1.Print "Ⅴ"
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.CurrentX = 50
Picture1.CurrentY = 0.024
Picture1.Print "f"

End Sub

Private Sub Command6_Click()
 On Error Resume Next
    Dim txttemp As Control
    For Each txttemp In Me
        txttemp.Text = ""
    Next
Picture1.Picture = LoadPicture("")


End Sub

Private Sub Command7_Click()
Text1.Text = 6#
Text2.Text = 3.8
Text3.Text = 2.5
Text4.Text = 1.6
Text5.Text = 1#
End Sub

Private Sub Command8_Click()
Picture1.Cls
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.DrawWidth = 1

Picture1.Scale (-10, 45)-(135, -4.7)

Picture1.Line (0, 0)-(125, 0)
Picture1.Line (0, 0)-(0, 42.5)
Picture1.CurrentX = 60
Picture1.CurrentY = -2.5
Picture1.Print "ua / (km / h)";
Picture1.CurrentX = -2
Picture1.CurrentY = 43.7
Picture1.Print "i(%)"
For I = 0 To 120 Step 20
If I <> 0 Then
Picture1.CurrentX = I
Picture1.CurrentY = 0.7
Picture1.Line (I, 0.7)-(I, 0)
End If
Next
For j = 0 To 120 Step 20
If j <> 0 Then
Picture1.CurrentX = j - 3
Picture1.CurrentY = -0.7
Picture1.Print j
Else
Picture1.CurrentX = -1
Picture1.CurrentY = -0.7
Picture1.Print 0
End If
Next
For K = 0 To 30 Step 10
If K <> 0 Then
Picture1.CurrentX = 1.3
Picture1.CurrentY = K
Picture1.Line (1.3, K)-(0, K)
End If
Next
For K = 0 To 30 Step 10
If K <> 0 Then
Picture1.CurrentX = -5
Picture1.CurrentY = K
Picture1.Print K
End If
Next

Dim n As Single, Ttq As Single, ua As Single, Ft As Single, a As Single, I1 As Double, I2 As Double, I3 As Double, I4 As Double, I5 As Double

For ua = 4 To 18 Step 0.01
I1 = Val(Text1)

n = ua * 5.83 / 0.367 / 0.377 * I1

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I1

D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, I * 100)
Next ua
For ua = 6 To 30 Step 0.01
I2 = Val(Text2)

n = ua * 5.83 / 0.367 / 0.377 * I2


Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I2
D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, I * 100)
Next ua
For ua = 12 To 42 Step 0.01
I3 = Val(Text3)


n = ua * 5.83 / 0.367 / 0.377 * I3

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I3
D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, I * 100)
Next ua

For ua = 22 To 70 Step 0.01
I4 = Val(Text4)

n = ua * 5.83 / 0.367 / 0.377 * I4

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I4

D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, I * 100)
Next ua
For ua = 25 To 101 Step 0.01
I5 = Val(Text5)

n = ua * 5.83 / 0.367 / 0.377 * I5

Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12

Ft = Ttq * 5.83 * 0.85 / 0.367 * I5
D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
Picture1.DrawWidth = 1.5
Picture1.ForeColor = RGB(100, 100, 225)
Picture1.PSet (ua, I * 100)
Next ua
For ua = 4 To 22 Step 0.01
n = ua * 5.83 / 0.367 / 0.377 * I1
Ttq = -19.313 + 295.27 * n / 1000 - 165.44 * n ^ 2 / 10 ^ 6 + 40.874 * n ^ 3 / 10 ^ 9 - 3.8445 * n ^ 4 / 10 ^ 12
Ft = Ttq * 5.83 * 0.85 / 0.367 * I1
D = (Ft - 2.77 / 21.15 * ua * ua - 3800 * 9.8 * 0.013) / 3800 / 9.8
B = Atn(D / Sqr(1 - D ^ 2))
I = Tan(B)
If I >= imax Then
imax = I
Else
imax = imax
End If
Next ua
Picture1.DrawWidth = 1
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (0, imax * 100)-(10, imax * 100)
Picture1.CurrentX = 5
Picture1.CurrentY = imax * 100 + 1
Picture1.Print "imax="; Int(imax * 100)
Picture1.CurrentX = 20
Picture1.CurrentY = 27
Picture1.Print "Ⅰ"
Picture1.CurrentX = 23
Picture1.CurrentY = 18
Picture1.Print "Ⅱ"
Picture1.CurrentX = 42
Picture1.CurrentY = 9.8
Picture1.Print "Ⅲ"
Picture1.CurrentX = 75
Picture1.CurrentY = 4
Picture1.Print "Ⅳ"
Picture1.CurrentX = 74
Picture1.CurrentY = 1.8
Picture1.Print "Ⅴ"

End Sub

Public Sub Form_Load()
'I1 = Val(Text1)
'I2 = Val(Text2)
'I3 = Val(Text3)
'I4 = Val(Text4)
'I5 = Val(Text5)


End Sub

