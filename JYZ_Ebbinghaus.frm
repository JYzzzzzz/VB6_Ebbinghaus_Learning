VERSION 5.00
Begin VB.Form JYZ_Ebbinghaus 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JYZ_Ebbinghaus_learning"
   ClientHeight    =   7080
   ClientLeft      =   4770
   ClientTop       =   2775
   ClientWidth     =   9525
   Icon            =   "JYZ_Ebbinghaus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9525
   Begin VB.TextBox txt_Days 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3020
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6615
      TabIndex        =   27
      Top             =   6345
      Width           =   2265
   End
   Begin VB.CommandButton cmd_ManualLvlChg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   6705
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmd_ManualLvlChg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmd_ManualLvlChg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmd_ManualLvlChg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2115
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0FF&
      Caption         =   "重置回本次刚开状态"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "其他操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   60
      TabIndex        =   16
      Top             =   1710
      Width           =   1935
      Begin VB.CheckBox chk_Priority_idx 
         BackColor       =   &H80000016&
         Caption         =   "复习优先级仅看序号"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   30
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton CmdRevisePlan 
         Caption         =   "显示今后复习计划"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   28
         Top             =   2115
         Width           =   1695
      End
      Begin VB.CommandButton cmdOneMore 
         Caption         =   "强行新学1知识点"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   21
         Top             =   1755
         Width           =   1695
      End
      Begin VB.CommandButton cmdLrnNum 
         Caption         =   "预设明天学习个数"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   19
         Top             =   1395
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加新知识点"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width           =   1695
      End
      Begin VB.CommandButton cmdChg 
         Caption         =   "修改该知识点"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   17
         Top             =   675
         Width           =   1695
      End
   End
   Begin VB.TextBox txtMessage 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   855
      Left            =   3180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "JYZ_Ebbinghaus.frx":030A
      Top             =   4500
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   300
      Top             =   6240
   End
   Begin VB.CommandButton cmdDontKnow 
      BackColor       =   &H0080C0FF&
      Caption         =   "模  糊"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1590
   End
   Begin VB.CommandButton cmdKnow 
      BackColor       =   &H0080FF80&
      Caption         =   "知  道"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1590
   End
   Begin VB.CommandButton cmdShowAnswer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "↓---- 显示答案"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2745
      Width           =   1815
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   2115
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "JYZ_Ebbinghaus.frx":0310
      Top             =   540
      Width           =   4935
   End
   Begin VB.CommandButton cmd1Next 
      BackColor       =   &H00FFC0C0&
      Caption         =   "开始下次学习"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmd1This 
      BackColor       =   &H00FFC0C0&
      Caption         =   "继续上次学习"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   540
      Width           =   1935
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2115
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "JYZ_Ebbinghaus.frx":034C
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label lblLV 
      BackColor       =   &H80000004&
      Caption         =   " LV:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   2835
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   4380
      X2              =   4380
      Y1              =   420
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6960
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label lblacquired 
      Caption         =   "基本掌握："
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lbllearned 
      Caption         =   "已学知识："
      Height          =   255
      Left            =   1500
      TabIndex        =   13
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblAll 
      Caption         =   "知识总数："
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2115
      Left            =   7155
      Top             =   60
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6120
      Left            =   7155
      Picture         =   "JYZ_Ebbinghaus.frx":0358
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "需新学："
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblNew 
      Caption         =   "0"
      Height          =   255
      Left            =   6540
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblReview 
      Caption         =   "0"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "需复习："
      Height          =   255
      Left            =   4620
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   " 问题：-----↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2835
      Width           =   1455
   End
End
Attribute VB_Name = "JYZ_Ebbinghaus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1             '数组下标从一开始计数

'------------------------------ 知识点存储数据，各参数
Dim Klg_Q() As String  'array: knowledge question
Dim Klg_A() As String  'array: knowledge answer
Dim Klg_L() As Integer  'array: knowledge level: 见下文
Dim Klg_D() As Integer  'array: knowledge days
Dim DataNum As Long     'data numbers

'------------------------------ 知识点等级全局量
Const LvlForget = -1
Const LvlNew = 0
Const Lvl1 = 1
Const Lvl2 = 2
Const Lvl3 = 5
Const Lvl4 = 8
Dim LvlMng() As Single     '记录level_manage.txt中的数据，用于等级的个性化管理
Dim LvlMng_DataSum As Byte   '数组LvlMng()的有效数据个数
Dim LvlMng_Multiple As Single  '倍数， 一般为1.5
Dim LvlMng_LevelMax As Integer '最大等级
Const Max_Lvl__revise_too_much = 32  '常量。当复习知识点过多时，用于计算第一优先级的最大分母
Dim ProLvl_ifKnow As Integer, ProLvl_ifDontKnow As Integer  '知识点下次等级的预判断

'------------------------------- 知识点下次复习天数全局量
Dim Days_Distribution() As Integer   '下次复习天数分布律
Dim Days_max_idx As Long          '最大的复习天数
Const DAYS_MAX_LIMIT = 4000       'Days_max_idx 最大不能大于 DAYS_MAX_LIMIT


'------------------------------- 按键是否可控变换位置汇总
Const FORM_INIT = 0
Const FORM_STUDY_QUESTION_ONLY = 1
Const FORM_STUDY_QUESTION_ANSWER = 2
Const FORM_STUDY_FINISH = 3
'1、form load      2、开始下次      3、继续上次
'4、学习完最后一个知识点



'存储在txt中的其他参数，目前：--------------------------------------------------
Dim LastDate As Date            '1、上次学习日期
Dim Study_Sum As Integer        '2、每日计划学习个数（可改）。
Dim Weight_TopRanking As Single '3 靠前的知识点的重要性权重
Dim Pic_Sum As Byte             '4、学习进度条显示图片总个数（可改）
Dim Pic_LastIndex As Byte       '5、上次学习进度条显示的图片的序号（不可改）。
Dim Days_of_NotLearnNewPoint As Integer  '6、连续没学新知识点的天数
Dim Data_Text_Idx As Integer        '7、多个data备份，记录idx


'当天学习管理，各参数--------------------------------------------
Dim Study_Index() As Integer      '存储知识点对应data序号
Dim Study_Whether() As Integer    '1 or 0表示是否已学
'Dim StudyNum As Integer     'studying numbers daily，即有几个知识点
Dim StudyNumReal As Integer 'studying numbers daily real，即最终需要点几下知道
Dim OrdinalNum As Integer   '表示今天学到第几个，<= StudyNumReal，即已经点了几下知道
Dim idx_newpoint As Long    '指向最新知识点的下标，当数据中无新知识点时，值为“Study_Sum+1”
Dim Flag_IfAdd As Byte      '是否有新增知识点标志位。1表示有。用于data备份

'-------------------- 其他辅助变量
Private Declare Function timeGetTime Lib "winmm.dll" () As Long   '随机数有关


'function: 手动改变知识点等级
Private Sub cmd_ManualLvlChg_Click(Index As Integer)

'------------------------------ 修改
Select Case Index   '控件数组操作实例
    Case 0        '-10
        ProLvl_ifKnow = ProLvl_ifKnow - 10
        ProLvl_ifDontKnow = ProLvl_ifDontKnow - 10
    Case 1        '-1
        ProLvl_ifKnow = ProLvl_ifKnow - 1
        ProLvl_ifDontKnow = ProLvl_ifDontKnow - 1
    Case 2        '+1
        ProLvl_ifKnow = ProLvl_ifKnow + 1
        ProLvl_ifDontKnow = ProLvl_ifDontKnow + 1
    Case 3        '+10
        ProLvl_ifKnow = ProLvl_ifKnow + 10
        ProLvl_ifDontKnow = ProLvl_ifDontKnow + 10
End Select
'------------------------------ 范围限制
If ProLvl_ifKnow <= 0 Then
    If Index = 0 Or Index = 1 Then
        ProLvl_ifKnow = -1   '-操作
    Else
        ProLvl_ifKnow = 1    '+操作
    End If
End If
If ProLvl_ifDontKnow <= 0 Then
    If Index = 0 Or Index = 1 Then
        ProLvl_ifDontKnow = -1   '-操作
    Else
        ProLvl_ifDontKnow = 1   '+操作
    End If
End If

cmdKnow.Caption = "知 道（" & ProLvl_ifKnow & ")"
cmdDontKnow.Caption = "模 糊（" & ProLvl_ifDontKnow & ")"
txt_Days.Visible = False

End Sub

'复习过多时计算第一优先级的函数
Private Function Priority_Cal(klg_i As Integer) As Double

Dim days As Double, level As Double

days = Klg_D(klg_i) - 1
level = Klg_L(klg_i)
If (level > Max_Lvl__revise_too_much) Then level = Max_Lvl__revise_too_much
                                 '设置一个最低优先级用于那些基本掌握的知识点
If (level < 1) Then level = 1    '防止level为0或负数
Priority_Cal = days / level

End Function

Private Sub cmd1Next_Click()

Dim i As Long
Dim TodayDate As Date
Dim date_diff As Integer
Dim Tmp_Int As Integer

'读取今日需复习的知识点------------------------------
TodayDate = Date
StudyNumReal = 0 '实际学习的知识点个数归零
For i = DataNum To 1 Step -1
    If Klg_L(i) <> 0 Then      '排除新知识点
        Klg_D(i) = Klg_D(i) - (TodayDate - LastDate)  '减去过了的天数
        If Klg_D(i) <= 0 Then       'judge and add to study array
            StudyNumReal = StudyNumReal + 1
            ReDim Preserve Study_Index(1 To StudyNumReal)
            ReDim Preserve Study_Whether(1 To StudyNumReal)
            Study_Index(StudyNumReal) = i     '存入知识点序号
            Study_Whether(StudyNumReal) = 0    '存入知识点今天是否学习状态，0，no；1，yes
            'Klg_D(i) = 1                       '将知识点的天数重新改为1
        End If
        
    End If
Next

'--------------------------------------------- 修改天数分布
date_diff = TodayDate - LastDate
i = 1
Do While i + date_diff <= Days_max_idx
    Days_Distribution(i) = Days_Distribution(i + date_diff)
    i = i + 1
Loop
Do While i <= Days_max_idx
    Days_Distribution(i) = 0
    i = i + 1
Loop
If CmdRevisePlan.Caption = "隐藏今后复习计划" Then CmdRevisePlan_Click


'确认今日学习个数------------------------------------
Dim s As String     '返回字符串
Dim note As String  '提示字符串
note = "-- 需复习的知识点数为：" & StudyNumReal
note = note & Chr(13) & Chr(10) & "-- 您已连续 " & Days_of_NotLearnNewPoint & " 次没有学习新知识点了"
note = note & Chr(13) & Chr(10) & "-- 今天您总共想学几个知识点？(取消则为默认值：" & Study_Sum & ")"
s = InputBox(note, "知识点学习数确认", Study_Sum)
If s <> "" Then
    Study_Sum = Val(s)
End If

'强迫症操作，将复习的知识点的显示顺序改为正序---------------
i = 1
Do While i < StudyNumReal + 1 - i
    Tmp_Int = Study_Index(i)
    Study_Index(i) = Study_Index(StudyNumReal + 1 - i)
    Study_Index(StudyNumReal + 1 - i) = Tmp_Int
    i = i + 1
Loop


'如果学习的个数小于复习的个数（需复习的过多）-------------------------------
Dim Num_front, Num_new As Integer  '计算空间，分别存储靠前知识点、最新知识点的个数
Dim Study_Index__mid As Integer     '计算空间，排序算法的中间变量
Dim j As Integer

If StudyNumReal > Study_Sum Then  '判断是否复习过多！！！！！
    Num_front = Int(Study_Sum * Weight_TopRanking)  '复习序号靠前的知识点的个数
    Num_new = Study_Sum - Num_front
    
    '----------------------------------------- 将知识点进行排序
    ' （优先级：(下次复习剩余天数-1)/知识点等级 、 序号）
    Dim Priority_lib As Double, Priority_insult As Double  '临时变量，存储第一优先级
    For i = 2 To (StudyNumReal - Num_new)  'i表示待比较知识点下标
        Study_Index__mid = Study_Index(i)
        
        Priority_insult = Priority_Cal(Study_Index__mid)
        If chk_Priority_idx.Value = 1 Then Priority_insult = 1  '不看剩余天数动态优先级，只看序号，用于长时间不学习后重新学习
        
        For j = i - 1 To 1 Step -1           'j表示已比较知识点下标
            Priority_lib = Priority_Cal(Study_Index(j))
            If Priority_insult < Priority_lib Then
                Study_Index(j + 1) = Study_Index(j)
            ElseIf Priority_insult = Priority_lib And Study_Index__mid < Study_Index(j) Then
                Study_Index(j + 1) = Study_Index(j)
            Else
                'Study_Index(j + 1) = Study_Index__mid
                Exit For
            End If
        Next
        Study_Index(j + 1) = Study_Index__mid
        
    Next
    
    
    '----------------------------------------- 剔除中间的多出的知识点
    For i = (Num_new - 1) To 0 Step -1
        Study_Index(Study_Sum - i) = Study_Index(StudyNumReal - i)
    Next
    
    '----------------------------------------- 对前“Study_Sum”个知识点排序（优先级：序号）
'    If Study_Sum > 1 Then
'    For i = 2 To Study_Sum                   'i表示待比较元素下标
'        Study_Index__mid = Study_Index(i)
'        For j = i - 1 To 1 Step -1           'j表示已比较元素下标
'            If Study_Index__mid < Study_Index(j) Then
'                Study_Index(j + 1) = Study_Index(j)
'            Else
'                'Study_Index(j + 1) = Study_Index__mid
'                Exit For
'            End If
'        Next
'        Study_Index(j + 1) = Study_Index__mid
'    Next
'    End If

    StudyNumReal = Study_Sum
    Days_of_NotLearnNewPoint = Days_of_NotLearnNewPoint + 1 '没学新知识点，天数加1
Else
    Days_of_NotLearnNewPoint = 0
End If
lblReview.Caption = StudyNumReal    '显示需复习个数


'------------------------------------------ 对前StudyNumReal个知识点进行乱序处理
Dim temp_double_array() As Double  'temperary array, store random value
Dim temp_double_mid As Double
If StudyNumReal > 1 Then
    ReDim Preserve temp_double_array(1 To StudyNumReal)
    '-------------------- 产生一个随机数组作为优先级
    Randomize
    For i = 1 To StudyNumReal
        temp_double_array(i) = Rnd * 100
    Next
    '-------------------- 对前StudyNumReal个知识点排序（优先级：temp_double_array数值由小到大）
    For i = 2 To StudyNumReal                   'i表示待比较元素下标
        Study_Index__mid = Study_Index(i)
        temp_double_mid = temp_double_array(i)
        For j = i - 1 To 1 Step -1           'j表示已比较元素下标
            If temp_double_mid < temp_double_array(j) Then
                Study_Index(j + 1) = Study_Index(j)
                temp_double_array(j + 1) = temp_double_array(j)
            Else
                'Study_Index(j + 1) = Study_Index__mid
                Exit For
            End If
        Next
        Study_Index(j + 1) = Study_Index__mid
        temp_double_array(j + 1) = temp_double_mid
    Next
End If
    

'添加新知识点---------------------------------------------
idx_newpoint = 1    '新知识点指针初始化
If StudyNumReal < Study_Sum Then
    ReDim Preserve Study_Index(1 To Study_Sum)
    ReDim Preserve Study_Whether(1 To Study_Sum)
    
    Do While (idx_newpoint <= DataNum And StudyNumReal < Study_Sum)   '重新遍历知识点（Level）数组
        If (Klg_L(idx_newpoint) = 0) Then       'if The Level is 0, put The Index into the list of points today
            StudyNumReal = StudyNumReal + 1
            Study_Index(StudyNumReal) = idx_newpoint
            Study_Whether(StudyNumReal) = 0
        End If
        idx_newpoint = idx_newpoint + 1
    Loop
End If

lblNew.Caption = StudyNumReal - lblReview.Caption

'完成上述操作后，仍然没有任何知识点-----------------------
If StudyNumReal = 0 Then
    OrdinalNum = 0
    ShowQuestion
End If
'other--------------------------------------------------
LastDate = TodayDate
'Open App.Path & "\Data\lastdate.txt" For Output As #4  '马上更新日期
'    Write #4, Date
'Close #4

'显示今天学习的第一个知识点-----------------------------
OrdinalNum = 0
ShowQuestion

End Sub

Private Sub cmd1This_Click()

'reading last study record--------------------------------------------
StudyNumReal = 0
OrdinalNum = 0
Open App.Path & "\Data\studying.txt" For Input As #3
Do While Not EOF(3)
    StudyNumReal = StudyNumReal + 1
    ReDim Preserve Study_Index(1 To StudyNumReal)
    ReDim Preserve Study_Whether(1 To StudyNumReal)
    Input #3, Study_Index(StudyNumReal), Study_Whether(StudyNumReal)
    If Study_Whether(StudyNumReal) = 1 Then
        OrdinalNum = OrdinalNum + 1
    Else
        'cmd1Next.Enabled = False
    End If
Loop
Close #3

'others----------------------------------
idx_newpoint = 1    '新知识点指针初始化
lblReview.Caption = StudyNumReal - OrdinalNum
ShowQuestion

End Sub

Private Sub cmdAdd_Click()

If txtQuestion.Text <> "" And txtAnswer.Text <> "" Then

    DataNum = DataNum + 1
    ReDim Preserve Klg_Q(1 To DataNum)
    ReDim Preserve Klg_A(1 To DataNum)
    ReDim Preserve Klg_L(1 To DataNum)
    ReDim Preserve Klg_D(1 To DataNum)
    Klg_Q(DataNum) = txtQuestion.Text
    Klg_A(DataNum) = txtAnswer.Text
    Klg_L(DataNum) = 0
    Klg_D(DataNum) = 0
    
    txtQuestion.Text = ""
    txtAnswer.Text = ""
    Flag_IfAdd = 1
    Call ShowMessage("Add succeed!", 2000)
Else
    Call ShowMessage("Can't add!", 3000)
End If

ShowDataState      '显示总知识数、已学知识等

End Sub

Private Sub cmdChg_Click()    '修改知识点
If txtQuestion.Text <> "" And txtAnswer.Text <> "" Then
    If Klg_Q(Study_Index(OrdinalNum)) <> txtQuestion.Text Or Klg_A(Study_Index(OrdinalNum)) <> txtAnswer.Text Then
        Klg_Q(Study_Index(OrdinalNum)) = txtQuestion.Text
        Klg_A(Study_Index(OrdinalNum)) = txtAnswer.Text
        Call ShowMessage("Change succeed!", 2000)
    End If
Else
    'change fail
    Call ShowMessage("Can't change!", 3000)
End If

End Sub

Private Sub cmdDontKnow_Click()

Study_Whether(OrdinalNum) = 1    '已学标志置1，不论是否“知道”

Klg_L(Study_Index(OrdinalNum)) = ProLvl_ifDontKnow
'Klg_D(Study_Index(OrdinalNum)) = Klg_L(Study_Index(OrdinalNum))
'更新等级与天数，即使有模糊知识的情况下未完成当天学习，也不会影响程序的正常运行。
'而模糊的知识点会在明天作为复习知识点

'------------------------------- 加入当天复习队列末尾
StudyNumReal = StudyNumReal + 1
ReDim Preserve Study_Index(1 To StudyNumReal)
ReDim Preserve Study_Whether(1 To StudyNumReal)
Study_Index(StudyNumReal) = Study_Index(OrdinalNum)
Study_Whether(StudyNumReal) = 0

cmdChg_Click    '自动检测知识点是否有修改
ShowQuestion    '显示下一问题

End Sub

Private Sub cmdKnow_Click()

Study_Whether(OrdinalNum) = 1    '已学标志置1，不论是否“知道”

'---------------------------------------------------------- 写入知识点等级
Klg_L(Study_Index(OrdinalNum)) = ProLvl_ifKnow

'---------------------------------------------------------- 写入天数
Klg_D(Study_Index(OrdinalNum)) = Days_Cal(ProLvl_ifKnow)
If ProLvl_ifKnow <= DAYS_MAX_LIMIT Then
    Days_Distribution(Klg_D(Study_Index(OrdinalNum))) = Days_Distribution(Klg_D(Study_Index(OrdinalNum))) + 1
End If

cmdChg_Click         '自动检测知识点是否有修改，如果有则自动修改
ShowQuestion

End Sub


Private Sub cmdKnow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If CmdRevisePlan.Caption = "显示今后复习计划" Then
    txt_Days.Visible = False
ElseIf CmdRevisePlan.Caption = "隐藏今后复习计划" Then
    If txt_Days.Visible = False Then
        txt_Days.Left = cmdKnow.Left + 50
        txt_Days.Top = cmdKnow.Top - 360
        txt_Days.Text = "天数：" & Days_Cal(ProLvl_ifKnow)
        txt_Days.Visible = True
    End If
End If

End Sub

'function: 计算下回复习天数
Private Function Days_Cal(next_lvl As Integer) As Integer

Dim days_candidate_array() As Integer   '天数候选空间，存放Days_Distribution数组下标
Dim days_revise_num_array() As Integer  '天数对应的已有的复习知识点个数的数组
Dim side_width As Integer
Dim candidate_num As Integer    '数组days_candidate_array的有效个数
Dim start_candi_day As Integer       '候选天数中最靠前的
Dim priority_reduce_step As Single   '当候选天数值大于知识点等级值时，优先级下降的步长

Dim i, best_i As Integer

If next_lvl * 2 > Days_max_idx And next_lvl <= DAYS_MAX_LIMIT Then  '扩大空间
    Days_max_idx = next_lvl * 2
    ReDim Preserve Days_Distribution(1 To Days_max_idx)
End If

If next_lvl > DAYS_MAX_LIMIT Then       '天数太大则不需要考虑天数的微调问题
    Days_Cal = next_lvl
    Exit Function
End If


side_width = CInt(next_lvl / 17.9)

If side_width = 0 Then
    Days_Cal = next_lvl
Else
    '------------------------------------------- 对side_width最大长度进行限制
    If (Study_Sum <= 30 And side_width > 8) Then side_width = 8
    If (Study_Sum > 30 And Study_Sum <= 60 And side_width > 12) Then side_width = 12
    '------------------------------------------- 开辟候选空间
    candidate_num = side_width + side_width + 1
    ReDim Preserve days_candidate_array(1 To candidate_num)
    ReDim Preserve days_revise_num_array(1 To candidate_num)
    start_candi_day = next_lvl - side_width
    For i = 1 To candidate_num
        days_candidate_array(i) = start_candi_day + i - 1
        days_revise_num_array(i) = Days_Distribution(days_candidate_array(i))
    Next
    '------------------------------------------- 当候选天数值大于知识点等级值时，优先级下降
    ' 优先级下降步长（以5+1+5为例）：+0,+0,+0,+0,+0,+0,+3,+3,+4,+4,+5
    priority_reduce_step = 3
    For i = side_width + 2 To candidate_num
        days_revise_num_array(i) = days_revise_num_array(i) + Int(priority_reduce_step)
        priority_reduce_step = priority_reduce_step + 0.5
    Next
    '------------------------------------------- 找到最优天数
    best_i = 1
    For i = 2 To candidate_num
        If days_revise_num_array(i) < days_revise_num_array(best_i) Then best_i = i
    Next
    Days_Cal = days_candidate_array(best_i)
    
End If

End Function


Private Sub cmdLrnNum_Click()

Dim i As Long
Dim Num_ReviseTomm As Integer   '存放明天需要复习的知识点个数
Dim ret_str As String

Num_ReviseTomm = 0
For i = 1 To DataNum
    If Klg_L(i) <> 0 And Klg_D(i) <= 1 Then   '
        Num_ReviseTomm = Num_ReviseTomm + 1
    End If
Next
ret_str = InputBox("明天需复习的知识点数为 (" & Num_ReviseTomm & ")  ,您总共想学几个？", "明天学习个数", Study_Sum)
If ret_str <> "" Then            '如果返回值不为空。表示InputBox中没有按Cancel或叉叉
    Study_Sum = ret_str
    Call ShowMessage("Change succeed!", 2000)
End If

End Sub

Private Sub cmdOneMore_Click()     '在今日学习中强行添加一个新知识点


Do While (idx_newpoint <= DataNum)   '继续遍历知识点（Level）数组
    If (Klg_L(idx_newpoint) = 0) Then       'if The Level is 0, put The Index into the list of points today
        StudyNumReal = StudyNumReal + 1
        ReDim Preserve Study_Index(1 To StudyNumReal)
        ReDim Preserve Study_Whether(1 To StudyNumReal)
        Study_Index(StudyNumReal) = idx_newpoint
        Study_Whether(StudyNumReal) = 0
        idx_newpoint = idx_newpoint + 1
        Exit Do
    Else
        idx_newpoint = idx_newpoint + 1
    End If
Loop

Days_of_NotLearnNewPoint = 0
'lblNew.Caption = StudyNumReal - lblReview.Caption
lblNew.Caption = Val(lblNew.Caption) + 1

End Sub

Private Sub cmdReset_Click()

Dim res As Integer

res = MsgBox("将重置回软件本次刚开状态（非本天刚开），是否继续？", vbOKCancel + vbExclamation, "重置提醒")
If res = vbOK Then
    FORM_LOAD
End If

End Sub

Private Sub CmdRevisePlan_Click()

Dim days As Integer        '下标
Dim revise_num As Integer
Dim Str As String

Dim tmp_single As Single   '临时空间

If CmdRevisePlan.Caption = "显示今后复习计划" Then
    List1.Clear
    List1.AddItem "今后复习计划"
    For days = 1 To Days_max_idx
        Str = days & ": "
        revise_num = Days_Distribution(days)
        Do While (revise_num > 0)
            Str = Str & "|"
            revise_num = revise_num - 1
        Loop
        Str = Str & " " & Days_Distribution(days)
        List1.AddItem Str
    Next
    List1.Top = Image1.Top
    List1.Left = Image1.Left
    List1.Height = Image1.Height
    List1.Width = Image1.Width
    List1.Visible = True
    CmdRevisePlan.Caption = "隐藏今后复习计划"
    
ElseIf CmdRevisePlan.Caption = "隐藏今后复习计划" Then
    List1.Visible = False
    CmdRevisePlan.Caption = "显示今后复习计划"
    
End If

End Sub

Private Sub cmdShowAnswer_Click()

Dim CurrentLvl As Integer   '知识点当前等级
Dim i As Byte
Dim pos As Integer    '子字符串在Klg_A()中的位置。用于重要知识点最大等级

'----------------------------------------------------------- 知识点等级预判断
CurrentLvl = Klg_L(Study_Index(OrdinalNum))  '获取知识点当前等级
lblLV.Caption = " LV: " & CurrentLvl       'show knowledge level
'---------------------------------------- 常量说明
'Const LvlForget = -1
'Const LvlNew = 0
'Const Lvl1 = 1
'Const Lvl2 = 2
'Const Lvl3 = 5
'Const Lvl4 = 8
'---------------------------------------- 如果知道
Select Case CurrentLvl
Case LvlForget
    ProLvl_ifKnow = LvlMng(1)
Case LvlNew
    ProLvl_ifKnow = LvlMng(4)
Case Is <= Int((LvlMng(LvlMng_DataSum) + LvlMng(LvlMng_DataSum - 1)) / 2#)
    For i = 2 To LvlMng_DataSum
        If CurrentLvl <= Int((LvlMng(i) + LvlMng(i - 1)) / 2#) Then
            '例如i=5，则 LvlMng(i)=12，LvlMng(i-1)=8，
            '((LvlMng(i) + LvlMng(i-1)) / 2#)=10
            '若CurrentLvl<=10，则预赋12，
            '若CurrentLvl=11，则在下一个循环计算，会被预赋18
            '一次赋值后，要跳出for循环防止再赋
            ProLvl_ifKnow = LvlMng(i)
            Exit For
        End If
    Next
Case Else     '超出level manage的最高等级上限
    ProLvl_ifKnow = CurrentLvl * LvlMng_Multiple
    If LvlMng_LevelMax <> 0 And ProLvl_ifKnow > LvlMng_LevelMax Then
        ProLvl_ifKnow = LvlMng_LevelMax
    End If
End Select
'------------- 若是重要知识点，限制最高等级
pos = InStr(Klg_A(Study_Index(OrdinalNum)), "(max_lvl:")
If pos > 0 And ProLvl_ifKnow > Val(Mid(Klg_A(Study_Index(OrdinalNum)), (pos + 9))) Then
    ProLvl_ifKnow = Val(Mid(Klg_A(Study_Index(OrdinalNum)), (pos + 9))) '数值提取
End If

'---------------------------------------- 如果模糊
Select Case CurrentLvl
Case Is <= 5
    ProLvl_ifDontKnow = LvlForget  '等级0~5时遗忘，1天后再次学习
Case Is <= 12
    ProLvl_ifDontKnow = LvlMng(1)       '等级5~12时遗忘，2天后再次学习
Case Is <= 20
    ProLvl_ifDontKnow = LvlMng(2)       '等级13~20时遗忘，5天后再次学习
Case Is <= 30
    ProLvl_ifDontKnow = LvlMng(3)       '等级21~30时遗忘，8天后再次学习
Case Else
    ProLvl_ifDontKnow = CurrentLvl / 5#   '整型与浮点运算时，自动类型转换
End Select

txtAnswer.Text = Klg_A(Study_Index(OrdinalNum))        'show answer
Call Cmd_EnableManage(FORM_STUDY_QUESTION_ANSWER)   'cmd enable manage，等级预改变显示

End Sub

Private Sub FORM_LOAD()
Dim i As Byte

'---------------------------------------- 确定窗口大小
JYZ_Ebbinghaus.Width = 9615
JYZ_Ebbinghaus.Height = 6670

'---------------------------------------- 初始化
Days_max_idx = 10
Max_Klg_Num__inDay = 0
ReDim Days_Distribution(1 To Days_max_idx)

'-------------------------------------------------- 读数据、参数

'------------------------------ read Other Parameters from otherparameter.txt
Open App.Path & "\Data\otherparameter.txt" For Input As #2
    Input #2, LastDate
    Input #2, Study_Sum
    Input #2, Weight_TopRanking
    Input #2, Pic_Sum
    Input #2, Pic_LastIndex
    Input #2, Days_of_NotLearnNewPoint
    Input #2, Data_Text_Idx
Close #2


'-------------------- read knowledge from data.txt
DataNum = 0
Open App.Path & "\Data\data_txt_group\data_v2__" & Data_Text_Idx & ".txt" For Input As #1
Do While Not EOF(1)
    DataNum = DataNum + 1
    ReDim Preserve Klg_Q(1 To DataNum)
    ReDim Preserve Klg_A(1 To DataNum)
    ReDim Preserve Klg_L(1 To DataNum)
    ReDim Preserve Klg_D(1 To DataNum)
    Input #1, var_check, Klg_Q(DataNum), Klg_A(DataNum), Klg_L(DataNum), Klg_D(DataNum)
    '修改艾宾浩斯天数----------------- for test
    'If Klg_L(DataNum) = 15 Then
    '    Klg_L(DataNum) = 12  '只改等级不改天数
    'ElseIf Klg_L(DataNum) = 30 Then
    '    Klg_L(DataNum) = 27
    'ElseIf Klg_L(DataNum) = 60 Then
    '    Klg_L(DataNum) = 27 * 1.5
    'End If
    '====================================
    If var_check <> "__check__" Or Klg_Q(DataNum) = "__check__" Or Klg_A(DataNum) = "__check__" Then
        MsgBox "data格式校验失败：" & vbCrLf & _
                "校验：" & var_check & vbCrLf & _
                "问题：" & Klg_Q(DataNum) & vbCrLf & _
                "回答：" & Klg_A(DataNum) & vbCrLf & _
                "（请不要关闭软件，先检查data，修改并做好备份，然后按重置按钮重新加载）"
        Exit Do
    End If
    
    
    '----------------------- add to Days Distribution
    If (Klg_D(DataNum) > 0) And (Klg_D(DataNum) < DAYS_MAX_LIMIT) Then
        If Klg_D(DataNum) > Days_max_idx Then
            Days_max_idx = Klg_D(DataNum)
            ReDim Preserve Days_Distribution(1 To Days_max_idx)  'space expand
        End If
        Days_Distribution(Klg_D(DataNum)) = Days_Distribution(Klg_D(DataNum)) + 1
    End If
Loop
Close #1

'-------------------- run once for changing the format of data.txt
'Open App.Path & "\Data\data_v2.txt" For Output As #1
'For n = 1 To DataNum
'    Write #1, "__check__", Klg_Q(n), Klg_A(n), Klg_L(n), Klg_D(n)
'Next
'Close #1


'-------------------- read from "level_manage.txt"
LvlMng_DataSum = 0
Open App.Path & "\Data\level_manage.txt" For Input As #2
Do While Not EOF(2)
    LvlMng_DataSum = LvlMng_DataSum + 1
    ReDim Preserve LvlMng(1 To LvlMng_DataSum)
    Input #2, LvlMng(LvlMng_DataSum)
Loop
Close #2
LvlMng_DataSum = LvlMng_DataSum - 2
LvlMng_Multiple = LvlMng(LvlMng_DataSum + 1)
LvlMng_LevelMax = LvlMng(LvlMng_DataSum + 2)
'================================================== 读数据、参数

'---------------------------------------- 按键使能管理
Call Cmd_EnableManage(FORM_INIT)

'---------------------------------------- 初始化所有标志位
Flag_IfAdd = 0

'---------------------------------------- 显示界面上的文本、图片
lblReview.Caption = 0
lblNew.Caption = 0
txtQuestion.Text = ""
txtAnswer.Text = ""
ShowDataState               '显示总知识数、已学知识等
If CmdRevisePlan.Caption = "隐藏今后复习计划" Then CmdRevisePlan_Click
txt_Days.Visible = False

' ---------------------------------------- picture initialization
Shape1.Height = Image1.Height
Shape1.Width = Image1.Width
ShowPicture

'---------------------------------------- 标志窗口启动结束
'Call ShowMessage("Hello, Owner!", 1000)
'Call Delayms(2000)
Call ShowMessage("Hello, Owner!", 3000)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Dim n As Long
Dim Str As String

'---------------------------------------- 保存忘记添加的新知识点（cmdADD键有效时有效）
If cmdAdd.Enabled = True Then   '-----【补充】属于误操作时的补救程序，可能出小bug
    cmdAdd_Click
End If

'---------------------------------------- save data

'-------------------- 保存参数
Data_Text_Idx = Data_Text_Idx + 1
If Data_Text_Idx > 9 Then
    Data_Text_Idx = 0
End If

Open App.Path & "\Data\otherparameter.txt" For Output As #2
    Write #2, LastDate
    Write #2, Study_Sum
    Write #2, Weight_TopRanking
    Write #2, Pic_Sum
    Write #2, Pic_LastIndex
    Write #2, Days_of_NotLearnNewPoint
    Write #2, Data_Text_Idx
Close #2

'-------------------- 保存知识点数据库数据
Open App.Path & "\Data\data_txt_group\data_v2__" & Data_Text_Idx & ".txt" For Output As #1
For n = 1 To DataNum
    Write #1, "__check__", Klg_Q(n), Klg_A(n), Klg_L(n), Klg_D(n)
Next
Close #1
Str = "-- 知识点数据库已保存到data_v2__" & Data_Text_Idx & ".txt"


'-------------------- 保存本次学习进度数据
If StudyNumReal > 0 Then    '如果实际学习个数大于0。按任何一个进入学习按钮将导致其>0
    Open App.Path & "\Data\studying.txt" For Output As #3
    For n = 1 To StudyNumReal
        Write #3, Study_Index(n), Study_Whether(n)
    Next
    'n = MsgBox("本次学习数据已保存至studying.txt", 0 + 64, "Goodbye")
    Close #3
    Str = Str & Chr(13) & Chr(10) & "-- 今日目前学习进度已保存至studying.txt"
End If

'-------------------- 保存知识点备份（2024-04：已添加一组data文件存储，backup淘汰）
'If Flag_IfAdd = 1 Then    '有新增知识点时保存一次，尽量减少更新次数
'    Open App.Path & "\Data\data_backup.txt" For Output As #4
'    For n = 1 To DataNum
'        Write #4, "__check__", Klg_Q(n), Klg_A(n), Klg_L(n), Klg_D(n)
'    Next
'    Close #4
'    Str = Str & Chr(13) & Chr(10) & "-- data_backup.txt已更新请检查！！"
'End If

'提示框-----------------------------------------------------------
MsgBox Str, 0 + 64, "Goodbye"

End Sub

Private Sub ShowQuestion()

'-------------------- show picture
If StudyNumReal > 0 Then
    Shape1.Height = Image1.Height / StudyNumReal * (StudyNumReal - OrdinalNum)
Else
    Shape1.Height = 0     'avoid bug when (StudyNumReal==0)
End If
'==================== show picture
lblLV.Caption = ""            '隐藏等级标签框
txt_Days.Visible = False       '隐藏天数文本框
If CmdRevisePlan.Caption = "隐藏今后复习计划" Then CmdRevisePlan_Click  '隐藏列表

OrdinalNum = OrdinalNum + 1
If OrdinalNum > StudyNumReal Then 'study finished
    txtQuestion.Text = "本次学习已完成，期待您下次学习。做任何事，坚持是最重要的！"
    txtAnswer.Text = ""
    Call Cmd_EnableManage(FORM_STUDY_FINISH)
    ShowDataState      '显示总知识数、已学知识等
Else                              'study not finished
    txtQuestion.Text = Klg_Q(Study_Index(OrdinalNum))
    txtAnswer.Text = ""
    Call Cmd_EnableManage(FORM_STUDY_QUESTION_ONLY)
End If

End Sub

Private Sub ShowDataState()
Dim i, LearnedNum, AcquiredNum As Long
LearnedNum = 0
AcquiredNum = 0
For i = 1 To DataNum
    If Klg_L(i) <> 0 Then LearnedNum = LearnedNum + 1
    If Klg_L(i) > 15 Then AcquiredNum = AcquiredNum + 1
    '即“基本掌握”是指记忆能达到12天的知识点
Next
lblAll.Caption = "知识总数：" & DataNum
lbllearned.Caption = "已学知识：" & LearnedNum
lblacquired.Caption = "基本掌握：" & AcquiredNum
End Sub

'function: 通用过程。显示一些提示信息，并启动定时器Timer1
Private Sub ShowMessage(Str As String, ms As Integer)
txtMessage.Text = Str
txtMessage.Visible = True
Timer1.Interval = ms
Timer1.Enabled = True

End Sub


Private Sub Label2_Click()

End Sub

'function: Timer1计时归零事件。信息框消失
Private Sub Timer1_Timer()
'Static t As Byte
Timer1.Enabled = False
txtMessage.Visible = False    '信息框变为不可见

End Sub

' 功能：获取某路径下所有文件名
' 参数：list(): 字符串数组，用于存放返回的文件名
'       sPath: 路径
'       Filter：后缀限定
' 输出：list()
Sub GetFileList(list() As String, ByVal sPath As String, ByVal Filter As String)
    '这是获取指定du文件夹下指定后缀名的文zhi件名称的过程，装入数组picname()中

Dim sDir As String
Dim sFilter() As String
Dim lngFilterIndex As Long
'Dim lngIndex As Long
Dim n As Integer

sFilter = Split(Filter, ",")
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

n = 0
For lngFilterIndex = LBound(sFilter) To UBound(sFilter)
    sDir = Dir(sPath & sFilter(lngFilterIndex))
    Do While Len(sDir) > 0
        n = n + 1
        ReDim Preserve list(1 To n)
        list(n) = sDir
        sDir = Dir
    Loop
Next

End Sub

'function: 右侧图片的随机选择过程。
'notes: this Sub must be after the "key management" Sub
Private Sub ShowPicture()

Dim rand_idx As Integer
Dim pic_list() As String

Randomize    'init random seed
Call GetFileList(pic_list, App.Path & "\pictures\", "*.jpg,*.png")  '获取路径下所有图片

' ------------------------------ 决策显示哪一张图片，并记录
If cmd1Next.Enabled = True Then  '若非当天初次进入学习，图片不变
    Do
        rand_idx = Int(Rnd * (UBound(pic_list) - 1 + 1)) + 1    '最后加1是因为图片从01开始命名
    Loop Until rand_idx <> Pic_LastIndex
    Pic_LastIndex = rand_idx
End If

Image1.Picture = LoadPicture(App.Path & "\Pictures\" & pic_list(Pic_LastIndex)) 'show picture

End Sub


'function: 基于软件状态的按键“使能、显示”的统一管理
Private Sub Cmd_EnableManage(which As Byte)
Dim i As Integer

'------------------------------ 回归按键常用状态
cmd1This.Enabled = False
cmd1Next.Enabled = False

cmdShowAnswer.Enabled = False
cmdKnow.Enabled = False
cmdKnow.Caption = "知  道"
cmdDontKnow.Enabled = False
cmdDontKnow.Caption = "模  糊"
For i = 0 To 3
    cmd_ManualLvlChg.Item(i).Enabled = False '控件数组属性操作实例
Next

cmdChg.Enabled = False
cmdAdd.Enabled = False
cmdLrnNum.Enabled = False   '不能操作按钮“预设明天学习个数”
cmdOneMore.Enabled = False  '不能操作“强行新学”按钮
chk_Priority_idx.Enabled = False


'------------------------------ 根据不同的界面调整按键状态
Select Case which
    Case FORM_INIT
        If Date - LastDate = 0 Then
            cmd1This.Enabled = True
        Else
            cmd1Next.Enabled = True
        End If
        cmdAdd.Enabled = True
        chk_Priority_idx.Enabled = True
    Case FORM_STUDY_QUESTION_ONLY
        cmdShowAnswer.Enabled = True
        cmdChg.Enabled = True
        cmdOneMore.Enabled = True
    Case FORM_STUDY_QUESTION_ANSWER
        cmdKnow.Enabled = True
        cmdKnow.Caption = "知 道（" & ProLvl_ifKnow & ")"
        cmdDontKnow.Enabled = True
        cmdDontKnow.Caption = "模 糊（" & ProLvl_ifDontKnow & ")"
        For i = 0 To 3
            cmd_ManualLvlChg.Item(i).Enabled = True
        Next
        cmdChg.Enabled = True
        cmdOneMore.Enabled = True
    Case FORM_STUDY_FINISH
        cmdAdd.Enabled = True
        cmdLrnNum.Enabled = True

End Select

End Sub

'delayms
Private Sub Delayms(ms As Long)
Dim Savetime As Double
Savetime = timeGetTime '记下开始时的时间
While timeGetTime < Savetime + ms '循环等待
    DoEvents '转让控制权，以便让操作系统处理其它的事件。
Wend
End Sub




