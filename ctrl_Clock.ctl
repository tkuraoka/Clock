VERSION 5.00
Begin VB.UserControl ctrl_Clock 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   ScaleHeight     =   1410
   ScaleWidth      =   2235
   Begin VB.PictureBox pct_Clock 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   720
      Picture         =   "ctrl_Clock.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.Timer Tick 
      Interval        =   1000
      Left            =   360
      Top             =   120
   End
   Begin VB.Label lbl_Time 
      Alignment       =   2  '中央揃え
      Caption         =   "12:00 am"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "ctrl_Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lastMinute As Integer        '時計に表示した最新の分
Private lastHour As Integer          '時計に表示した最新の時間

Private lastX As Integer             '以前の秒針の終点
Private lastY As Integer

Private Sub Tick_Timer()
    Const pi = 3.141592653           'PIの値を定義する
    Dim t                            '時刻情報
    Dim x As Integer                 'lastXと同じ型を使用する
    t = Now
    sec = Second(t)
    Min = Minute(t)
    hr = Hour(t)
    
    pct_Clock.Scale (-16, 16)-(16, -16)        '時計アイコンのスケールを設定する
    '
    '分が変わっていればキャプションを更新し、
    'そのあとで、すべての針を消し、描画し直す
    '
    
    If Min <> lastMinute Or hr <> lastHour Then
        lbl_Time.Caption = Format$(t, "h::mm AM/PM")
        lastMinute = Min           '新しい現在時刻を変数に格納する
        lastHour = hr
        pct_Clock.Cls
        lastX = 999                '秒針が存在しないことを示す
        
        pct_Clock.DrawWidth = 2
        pct_Clock.DrawMode = 13    '消去不能な描画モードに変更する
        
        h = hr + Min / 60
        x = 5 * Sin(h * pi / 6)    '時計の終点
        y = 5 * Cos(h * pi / 6)
        pct_Clock.Line (0, 0)-(x, y)
        pct_Clock.DrawWidth = 1    '線の太さを1ピクセルに戻す
        
        x = 8 * Sin(Min * pi / 30)     '分針の終点
        y = 8 * Cos(Min * pi / 30)
        pct_Clock.Line (0, 0)-(x, y)
        pct_Clock.DrawWidth = 1
    End If
    
    pct_Clock.DrawMode = 10        '消去可能な描画モードに変更する
    red = RGB(255, 0, 0)           '赤い色を定義する
    
    x = 10 * Sin(sec * pi / 30)             '秒針の終点を計算する
    y = 10 * Cos(sec * pi / 30)
        
    If lastX <> 999 Then                                       '以前の秒針を消去する
        pct_Clock.Line (0, 0)-(lastX, lastY), red           '秒針を描く
    End If
    pct_Clock.Line (0, 0)-(x, y), red
    
    lastX = x
    lastY = y
End Sub

Private Sub UserControl_Initialize()
    lastX = 999     '秒針が存在しないことを示す
    'Form1.WindowState = 1       'ウィンドウを最小化する
End Sub
