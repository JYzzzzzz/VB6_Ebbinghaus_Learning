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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "���ûر��θտ�״̬"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��ϰ���ȼ��������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��ʾ���ϰ�ƻ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "ǿ����ѧ1֪ʶ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "Ԥ������ѧϰ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�����֪ʶ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�޸ĸ�֪ʶ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "ģ  ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "֪  ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��---- ��ʾ��"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "΢���ź�"
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
      Caption         =   "��ʼ�´�ѧϰ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����ϴ�ѧϰ"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "΢���ź�"
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
         Name            =   "����"
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
      Caption         =   "�������գ�"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lbllearned 
      Caption         =   "��ѧ֪ʶ��"
      Height          =   255
      Left            =   1500
      TabIndex        =   13
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblAll 
      Caption         =   "֪ʶ������"
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
      Caption         =   "����ѧ��"
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
      Caption         =   "�踴ϰ��"
      Height          =   255
      Left            =   4620
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   " ���⣺-----��"
      BeginProperty Font 
         Name            =   "����"
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
Option Base 1             '�����±��һ��ʼ����

'------------------------------ ֪ʶ��洢���ݣ�������
Dim Klg_Q() As String  'array: knowledge question
Dim Klg_A() As String  'array: knowledge answer
Dim Klg_L() As Integer  'array: knowledge level: ������
Dim Klg_D() As Integer  'array: knowledge days
Dim DataNum As Long     'data numbers

'------------------------------ ֪ʶ��ȼ�ȫ����
Const LvlForget = -1
Const LvlNew = 0
Const Lvl1 = 1
Const Lvl2 = 2
Const Lvl3 = 5
Const Lvl4 = 8
Dim LvlMng() As Single     '��¼level_manage.txt�е����ݣ����ڵȼ��ĸ��Ի�����
Dim LvlMng_DataSum As Byte   '����LvlMng()����Ч���ݸ���
Dim LvlMng_Multiple As Single  '������ һ��Ϊ1.5
Dim LvlMng_LevelMax As Integer '���ȼ�
Const Max_Lvl__revise_too_much = 32  '����������ϰ֪ʶ�����ʱ�����ڼ����һ���ȼ�������ĸ
Dim ProLvl_ifKnow As Integer, ProLvl_ifDontKnow As Integer  '֪ʶ���´εȼ���Ԥ�ж�

'------------------------------- ֪ʶ���´θ�ϰ����ȫ����
Dim Days_Distribution() As Integer   '�´θ�ϰ�����ֲ���
Dim Days_max_idx As Long          '���ĸ�ϰ����
Const DAYS_MAX_LIMIT = 4000       'Days_max_idx ����ܴ��� DAYS_MAX_LIMIT


'------------------------------- �����Ƿ�ɿر任λ�û���
Const FORM_INIT = 0
Const FORM_STUDY_QUESTION_ONLY = 1
Const FORM_STUDY_QUESTION_ANSWER = 2
Const FORM_STUDY_FINISH = 3
'1��form load      2����ʼ�´�      3�������ϴ�
'4��ѧϰ�����һ��֪ʶ��



'�洢��txt�е�����������Ŀǰ��--------------------------------------------------
Dim LastDate As Date            '1���ϴ�ѧϰ����
Dim Study_Sum As Integer        '2��ÿ�ռƻ�ѧϰ�������ɸģ���
Dim Weight_TopRanking As Single '3 ��ǰ��֪ʶ�����Ҫ��Ȩ��
Dim Pic_Sum As Byte             '4��ѧϰ��������ʾͼƬ�ܸ������ɸģ�
Dim Pic_LastIndex As Byte       '5���ϴ�ѧϰ��������ʾ��ͼƬ����ţ����ɸģ���
Dim Days_of_NotLearnNewPoint As Integer  '6������ûѧ��֪ʶ�������
Dim Data_Text_Idx As Integer        '7�����data���ݣ���¼idx


'����ѧϰ����������--------------------------------------------
Dim Study_Index() As Integer      '�洢֪ʶ���Ӧdata���
Dim Study_Whether() As Integer    '1 or 0��ʾ�Ƿ���ѧ
'Dim StudyNum As Integer     'studying numbers daily�����м���֪ʶ��
Dim StudyNumReal As Integer 'studying numbers daily real����������Ҫ�㼸��֪��
Dim OrdinalNum As Integer   '��ʾ����ѧ���ڼ�����<= StudyNumReal�����Ѿ����˼���֪��
Dim idx_newpoint As Long    'ָ������֪ʶ����±꣬������������֪ʶ��ʱ��ֵΪ��Study_Sum+1��
Dim Flag_IfAdd As Byte      '�Ƿ�������֪ʶ���־λ��1��ʾ�С�����data����

'-------------------- ������������
Private Declare Function timeGetTime Lib "winmm.dll" () As Long   '������й�


'function: �ֶ��ı�֪ʶ��ȼ�
Private Sub cmd_ManualLvlChg_Click(Index As Integer)

'------------------------------ �޸�
Select Case Index   '�ؼ��������ʵ��
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
'------------------------------ ��Χ����
If ProLvl_ifKnow <= 0 Then
    If Index = 0 Or Index = 1 Then
        ProLvl_ifKnow = -1   '-����
    Else
        ProLvl_ifKnow = 1    '+����
    End If
End If
If ProLvl_ifDontKnow <= 0 Then
    If Index = 0 Or Index = 1 Then
        ProLvl_ifDontKnow = -1   '-����
    Else
        ProLvl_ifDontKnow = 1   '+����
    End If
End If

cmdKnow.Caption = "֪ ����" & ProLvl_ifKnow & ")"
cmdDontKnow.Caption = "ģ ����" & ProLvl_ifDontKnow & ")"
txt_Days.Visible = False

End Sub

'��ϰ����ʱ�����һ���ȼ��ĺ���
Private Function Priority_Cal(klg_i As Integer) As Double

Dim days As Double, level As Double

days = Klg_D(klg_i) - 1
level = Klg_L(klg_i)
If (level > Max_Lvl__revise_too_much) Then level = Max_Lvl__revise_too_much
                                 '����һ��������ȼ�������Щ�������յ�֪ʶ��
If (level < 1) Then level = 1    '��ֹlevelΪ0����
Priority_Cal = days / level

End Function

Private Sub cmd1Next_Click()

Dim i As Long
Dim TodayDate As Date
Dim date_diff As Integer
Dim Tmp_Int As Integer

'��ȡ�����踴ϰ��֪ʶ��------------------------------
TodayDate = Date
StudyNumReal = 0 'ʵ��ѧϰ��֪ʶ���������
For i = DataNum To 1 Step -1
    If Klg_L(i) <> 0 Then      '�ų���֪ʶ��
        Klg_D(i) = Klg_D(i) - (TodayDate - LastDate)  '��ȥ���˵�����
        If Klg_D(i) <= 0 Then       'judge and add to study array
            StudyNumReal = StudyNumReal + 1
            ReDim Preserve Study_Index(1 To StudyNumReal)
            ReDim Preserve Study_Whether(1 To StudyNumReal)
            Study_Index(StudyNumReal) = i     '����֪ʶ�����
            Study_Whether(StudyNumReal) = 0    '����֪ʶ������Ƿ�ѧϰ״̬��0��no��1��yes
            'Klg_D(i) = 1                       '��֪ʶ����������¸�Ϊ1
        End If
        
    End If
Next

'--------------------------------------------- �޸������ֲ�
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
If CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�" Then CmdRevisePlan_Click


'ȷ�Ͻ���ѧϰ����------------------------------------
Dim s As String     '�����ַ���
Dim note As String  '��ʾ�ַ���
note = "-- �踴ϰ��֪ʶ����Ϊ��" & StudyNumReal
note = note & Chr(13) & Chr(10) & "-- �������� " & Days_of_NotLearnNewPoint & " ��û��ѧϰ��֪ʶ����"
note = note & Chr(13) & Chr(10) & "-- �������ܹ���ѧ����֪ʶ�㣿(ȡ����ΪĬ��ֵ��" & Study_Sum & ")"
s = InputBox(note, "֪ʶ��ѧϰ��ȷ��", Study_Sum)
If s <> "" Then
    Study_Sum = Val(s)
End If

'ǿ��֢����������ϰ��֪ʶ�����ʾ˳���Ϊ����---------------
i = 1
Do While i < StudyNumReal + 1 - i
    Tmp_Int = Study_Index(i)
    Study_Index(i) = Study_Index(StudyNumReal + 1 - i)
    Study_Index(StudyNumReal + 1 - i) = Tmp_Int
    i = i + 1
Loop


'���ѧϰ�ĸ���С�ڸ�ϰ�ĸ������踴ϰ�Ĺ��ࣩ-------------------------------
Dim Num_front, Num_new As Integer  '����ռ䣬�ֱ�洢��ǰ֪ʶ�㡢����֪ʶ��ĸ���
Dim Study_Index__mid As Integer     '����ռ䣬�����㷨���м����
Dim j As Integer

If StudyNumReal > Study_Sum Then  '�ж��Ƿ�ϰ���࣡��������
    Num_front = Int(Study_Sum * Weight_TopRanking)  '��ϰ��ſ�ǰ��֪ʶ��ĸ���
    Num_new = Study_Sum - Num_front
    
    '----------------------------------------- ��֪ʶ���������
    ' �����ȼ���(�´θ�ϰʣ������-1)/֪ʶ��ȼ� �� ��ţ�
    Dim Priority_lib As Double, Priority_insult As Double  '��ʱ�������洢��һ���ȼ�
    For i = 2 To (StudyNumReal - Num_new)  'i��ʾ���Ƚ�֪ʶ���±�
        Study_Index__mid = Study_Index(i)
        
        Priority_insult = Priority_Cal(Study_Index__mid)
        If chk_Priority_idx.Value = 1 Then Priority_insult = 1  '����ʣ��������̬���ȼ���ֻ����ţ����ڳ�ʱ�䲻ѧϰ������ѧϰ
        
        For j = i - 1 To 1 Step -1           'j��ʾ�ѱȽ�֪ʶ���±�
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
    
    
    '----------------------------------------- �޳��м�Ķ����֪ʶ��
    For i = (Num_new - 1) To 0 Step -1
        Study_Index(Study_Sum - i) = Study_Index(StudyNumReal - i)
    Next
    
    '----------------------------------------- ��ǰ��Study_Sum����֪ʶ���������ȼ�����ţ�
'    If Study_Sum > 1 Then
'    For i = 2 To Study_Sum                   'i��ʾ���Ƚ�Ԫ���±�
'        Study_Index__mid = Study_Index(i)
'        For j = i - 1 To 1 Step -1           'j��ʾ�ѱȽ�Ԫ���±�
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
    Days_of_NotLearnNewPoint = Days_of_NotLearnNewPoint + 1 'ûѧ��֪ʶ�㣬������1
Else
    Days_of_NotLearnNewPoint = 0
End If
lblReview.Caption = StudyNumReal    '��ʾ�踴ϰ����


'------------------------------------------ ��ǰStudyNumReal��֪ʶ�����������
Dim temp_double_array() As Double  'temperary array, store random value
Dim temp_double_mid As Double
If StudyNumReal > 1 Then
    ReDim Preserve temp_double_array(1 To StudyNumReal)
    '-------------------- ����һ�����������Ϊ���ȼ�
    Randomize
    For i = 1 To StudyNumReal
        temp_double_array(i) = Rnd * 100
    Next
    '-------------------- ��ǰStudyNumReal��֪ʶ���������ȼ���temp_double_array��ֵ��С����
    For i = 2 To StudyNumReal                   'i��ʾ���Ƚ�Ԫ���±�
        Study_Index__mid = Study_Index(i)
        temp_double_mid = temp_double_array(i)
        For j = i - 1 To 1 Step -1           'j��ʾ�ѱȽ�Ԫ���±�
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
    

'�����֪ʶ��---------------------------------------------
idx_newpoint = 1    '��֪ʶ��ָ���ʼ��
If StudyNumReal < Study_Sum Then
    ReDim Preserve Study_Index(1 To Study_Sum)
    ReDim Preserve Study_Whether(1 To Study_Sum)
    
    Do While (idx_newpoint <= DataNum And StudyNumReal < Study_Sum)   '���±���֪ʶ�㣨Level������
        If (Klg_L(idx_newpoint) = 0) Then       'if The Level is 0, put The Index into the list of points today
            StudyNumReal = StudyNumReal + 1
            Study_Index(StudyNumReal) = idx_newpoint
            Study_Whether(StudyNumReal) = 0
        End If
        idx_newpoint = idx_newpoint + 1
    Loop
End If

lblNew.Caption = StudyNumReal - lblReview.Caption

'���������������Ȼû���κ�֪ʶ��-----------------------
If StudyNumReal = 0 Then
    OrdinalNum = 0
    ShowQuestion
End If
'other--------------------------------------------------
LastDate = TodayDate
'Open App.Path & "\Data\lastdate.txt" For Output As #4  '���ϸ�������
'    Write #4, Date
'Close #4

'��ʾ����ѧϰ�ĵ�һ��֪ʶ��-----------------------------
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
idx_newpoint = 1    '��֪ʶ��ָ���ʼ��
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

ShowDataState      '��ʾ��֪ʶ������ѧ֪ʶ��

End Sub

Private Sub cmdChg_Click()    '�޸�֪ʶ��
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

Study_Whether(OrdinalNum) = 1    '��ѧ��־��1�������Ƿ�֪����

Klg_L(Study_Index(OrdinalNum)) = ProLvl_ifDontKnow
'Klg_D(Study_Index(OrdinalNum)) = Klg_L(Study_Index(OrdinalNum))
'���µȼ�����������ʹ��ģ��֪ʶ�������δ��ɵ���ѧϰ��Ҳ����Ӱ�������������С�
'��ģ����֪ʶ�����������Ϊ��ϰ֪ʶ��

'------------------------------- ���뵱�츴ϰ����ĩβ
StudyNumReal = StudyNumReal + 1
ReDim Preserve Study_Index(1 To StudyNumReal)
ReDim Preserve Study_Whether(1 To StudyNumReal)
Study_Index(StudyNumReal) = Study_Index(OrdinalNum)
Study_Whether(StudyNumReal) = 0

cmdChg_Click    '�Զ����֪ʶ���Ƿ����޸�
ShowQuestion    '��ʾ��һ����

End Sub

Private Sub cmdKnow_Click()

Study_Whether(OrdinalNum) = 1    '��ѧ��־��1�������Ƿ�֪����

'---------------------------------------------------------- д��֪ʶ��ȼ�
Klg_L(Study_Index(OrdinalNum)) = ProLvl_ifKnow

'---------------------------------------------------------- д������
Klg_D(Study_Index(OrdinalNum)) = Days_Cal(ProLvl_ifKnow)
If ProLvl_ifKnow <= DAYS_MAX_LIMIT Then
    Days_Distribution(Klg_D(Study_Index(OrdinalNum))) = Days_Distribution(Klg_D(Study_Index(OrdinalNum))) + 1
End If

cmdChg_Click         '�Զ����֪ʶ���Ƿ����޸ģ���������Զ��޸�
ShowQuestion

End Sub


Private Sub cmdKnow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If CmdRevisePlan.Caption = "��ʾ���ϰ�ƻ�" Then
    txt_Days.Visible = False
ElseIf CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�" Then
    If txt_Days.Visible = False Then
        txt_Days.Left = cmdKnow.Left + 50
        txt_Days.Top = cmdKnow.Top - 360
        txt_Days.Text = "������" & Days_Cal(ProLvl_ifKnow)
        txt_Days.Visible = True
    End If
End If

End Sub

'function: �����»ظ�ϰ����
Private Function Days_Cal(next_lvl As Integer) As Integer

Dim days_candidate_array() As Integer   '������ѡ�ռ䣬���Days_Distribution�����±�
Dim days_revise_num_array() As Integer  '������Ӧ�����еĸ�ϰ֪ʶ�����������
Dim side_width As Integer
Dim candidate_num As Integer    '����days_candidate_array����Ч����
Dim start_candi_day As Integer       '��ѡ�������ǰ��
Dim priority_reduce_step As Single   '����ѡ����ֵ����֪ʶ��ȼ�ֵʱ�����ȼ��½��Ĳ���

Dim i, best_i As Integer

If next_lvl * 2 > Days_max_idx And next_lvl <= DAYS_MAX_LIMIT Then  '����ռ�
    Days_max_idx = next_lvl * 2
    ReDim Preserve Days_Distribution(1 To Days_max_idx)
End If

If next_lvl > DAYS_MAX_LIMIT Then       '����̫������Ҫ����������΢������
    Days_Cal = next_lvl
    Exit Function
End If


side_width = CInt(next_lvl / 17.9)

If side_width = 0 Then
    Days_Cal = next_lvl
Else
    '------------------------------------------- ��side_width��󳤶Ƚ�������
    If (Study_Sum <= 30 And side_width > 8) Then side_width = 8
    If (Study_Sum > 30 And Study_Sum <= 60 And side_width > 12) Then side_width = 12
    '------------------------------------------- ���ٺ�ѡ�ռ�
    candidate_num = side_width + side_width + 1
    ReDim Preserve days_candidate_array(1 To candidate_num)
    ReDim Preserve days_revise_num_array(1 To candidate_num)
    start_candi_day = next_lvl - side_width
    For i = 1 To candidate_num
        days_candidate_array(i) = start_candi_day + i - 1
        days_revise_num_array(i) = Days_Distribution(days_candidate_array(i))
    Next
    '------------------------------------------- ����ѡ����ֵ����֪ʶ��ȼ�ֵʱ�����ȼ��½�
    ' ���ȼ��½���������5+1+5Ϊ������+0,+0,+0,+0,+0,+0,+3,+3,+4,+4,+5
    priority_reduce_step = 3
    For i = side_width + 2 To candidate_num
        days_revise_num_array(i) = days_revise_num_array(i) + Int(priority_reduce_step)
        priority_reduce_step = priority_reduce_step + 0.5
    Next
    '------------------------------------------- �ҵ���������
    best_i = 1
    For i = 2 To candidate_num
        If days_revise_num_array(i) < days_revise_num_array(best_i) Then best_i = i
    Next
    Days_Cal = days_candidate_array(best_i)
    
End If

End Function


Private Sub cmdLrnNum_Click()

Dim i As Long
Dim Num_ReviseTomm As Integer   '���������Ҫ��ϰ��֪ʶ�����
Dim ret_str As String

Num_ReviseTomm = 0
For i = 1 To DataNum
    If Klg_L(i) <> 0 And Klg_D(i) <= 1 Then   '
        Num_ReviseTomm = Num_ReviseTomm + 1
    End If
Next
ret_str = InputBox("�����踴ϰ��֪ʶ����Ϊ (" & Num_ReviseTomm & ")  ,���ܹ���ѧ������", "����ѧϰ����", Study_Sum)
If ret_str <> "" Then            '�������ֵ��Ϊ�ա���ʾInputBox��û�а�Cancel����
    Study_Sum = ret_str
    Call ShowMessage("Change succeed!", 2000)
End If

End Sub

Private Sub cmdOneMore_Click()     '�ڽ���ѧϰ��ǿ�����һ����֪ʶ��


Do While (idx_newpoint <= DataNum)   '��������֪ʶ�㣨Level������
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

res = MsgBox("�����û�������θտ�״̬���Ǳ���տ������Ƿ������", vbOKCancel + vbExclamation, "��������")
If res = vbOK Then
    FORM_LOAD
End If

End Sub

Private Sub CmdRevisePlan_Click()

Dim days As Integer        '�±�
Dim revise_num As Integer
Dim Str As String

Dim tmp_single As Single   '��ʱ�ռ�

If CmdRevisePlan.Caption = "��ʾ���ϰ�ƻ�" Then
    List1.Clear
    List1.AddItem "���ϰ�ƻ�"
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
    CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�"
    
ElseIf CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�" Then
    List1.Visible = False
    CmdRevisePlan.Caption = "��ʾ���ϰ�ƻ�"
    
End If

End Sub

Private Sub cmdShowAnswer_Click()

Dim CurrentLvl As Integer   '֪ʶ�㵱ǰ�ȼ�
Dim i As Byte
Dim pos As Integer    '���ַ�����Klg_A()�е�λ�á�������Ҫ֪ʶ�����ȼ�

'----------------------------------------------------------- ֪ʶ��ȼ�Ԥ�ж�
CurrentLvl = Klg_L(Study_Index(OrdinalNum))  '��ȡ֪ʶ�㵱ǰ�ȼ�
lblLV.Caption = " LV: " & CurrentLvl       'show knowledge level
'---------------------------------------- ����˵��
'Const LvlForget = -1
'Const LvlNew = 0
'Const Lvl1 = 1
'Const Lvl2 = 2
'Const Lvl3 = 5
'Const Lvl4 = 8
'---------------------------------------- ���֪��
Select Case CurrentLvl
Case LvlForget
    ProLvl_ifKnow = LvlMng(1)
Case LvlNew
    ProLvl_ifKnow = LvlMng(4)
Case Is <= Int((LvlMng(LvlMng_DataSum) + LvlMng(LvlMng_DataSum - 1)) / 2#)
    For i = 2 To LvlMng_DataSum
        If CurrentLvl <= Int((LvlMng(i) + LvlMng(i - 1)) / 2#) Then
            '����i=5���� LvlMng(i)=12��LvlMng(i-1)=8��
            '((LvlMng(i) + LvlMng(i-1)) / 2#)=10
            '��CurrentLvl<=10����Ԥ��12��
            '��CurrentLvl=11��������һ��ѭ�����㣬�ᱻԤ��18
            'һ�θ�ֵ��Ҫ����forѭ����ֹ�ٸ�
            ProLvl_ifKnow = LvlMng(i)
            Exit For
        End If
    Next
Case Else     '����level manage����ߵȼ�����
    ProLvl_ifKnow = CurrentLvl * LvlMng_Multiple
    If LvlMng_LevelMax <> 0 And ProLvl_ifKnow > LvlMng_LevelMax Then
        ProLvl_ifKnow = LvlMng_LevelMax
    End If
End Select
'------------- ������Ҫ֪ʶ�㣬������ߵȼ�
pos = InStr(Klg_A(Study_Index(OrdinalNum)), "(max_lvl:")
If pos > 0 And ProLvl_ifKnow > Val(Mid(Klg_A(Study_Index(OrdinalNum)), (pos + 9))) Then
    ProLvl_ifKnow = Val(Mid(Klg_A(Study_Index(OrdinalNum)), (pos + 9))) '��ֵ��ȡ
End If

'---------------------------------------- ���ģ��
Select Case CurrentLvl
Case Is <= 5
    ProLvl_ifDontKnow = LvlForget  '�ȼ�0~5ʱ������1����ٴ�ѧϰ
Case Is <= 12
    ProLvl_ifDontKnow = LvlMng(1)       '�ȼ�5~12ʱ������2����ٴ�ѧϰ
Case Is <= 20
    ProLvl_ifDontKnow = LvlMng(2)       '�ȼ�13~20ʱ������5����ٴ�ѧϰ
Case Is <= 30
    ProLvl_ifDontKnow = LvlMng(3)       '�ȼ�21~30ʱ������8����ٴ�ѧϰ
Case Else
    ProLvl_ifDontKnow = CurrentLvl / 5#   '�����븡������ʱ���Զ�����ת��
End Select

txtAnswer.Text = Klg_A(Study_Index(OrdinalNum))        'show answer
Call Cmd_EnableManage(FORM_STUDY_QUESTION_ANSWER)   'cmd enable manage���ȼ�Ԥ�ı���ʾ

End Sub

Private Sub FORM_LOAD()
Dim i As Byte

'---------------------------------------- ȷ�����ڴ�С
JYZ_Ebbinghaus.Width = 9615
JYZ_Ebbinghaus.Height = 6670

'---------------------------------------- ��ʼ��
Days_max_idx = 10
Max_Klg_Num__inDay = 0
ReDim Days_Distribution(1 To Days_max_idx)

'-------------------------------------------------- �����ݡ�����

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
    '�޸İ�����˹����----------------- for test
    'If Klg_L(DataNum) = 15 Then
    '    Klg_L(DataNum) = 12  'ֻ�ĵȼ���������
    'ElseIf Klg_L(DataNum) = 30 Then
    '    Klg_L(DataNum) = 27
    'ElseIf Klg_L(DataNum) = 60 Then
    '    Klg_L(DataNum) = 27 * 1.5
    'End If
    '====================================
    If var_check <> "__check__" Or Klg_Q(DataNum) = "__check__" Or Klg_A(DataNum) = "__check__" Then
        MsgBox "data��ʽУ��ʧ�ܣ�" & vbCrLf & _
                "У�飺" & var_check & vbCrLf & _
                "���⣺" & Klg_Q(DataNum) & vbCrLf & _
                "�ش�" & Klg_A(DataNum) & vbCrLf & _
                "���벻Ҫ�ر�������ȼ��data���޸Ĳ����ñ��ݣ�Ȼ�����ð�ť���¼��أ�"
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
'================================================== �����ݡ�����

'---------------------------------------- ����ʹ�ܹ���
Call Cmd_EnableManage(FORM_INIT)

'---------------------------------------- ��ʼ�����б�־λ
Flag_IfAdd = 0

'---------------------------------------- ��ʾ�����ϵ��ı���ͼƬ
lblReview.Caption = 0
lblNew.Caption = 0
txtQuestion.Text = ""
txtAnswer.Text = ""
ShowDataState               '��ʾ��֪ʶ������ѧ֪ʶ��
If CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�" Then CmdRevisePlan_Click
txt_Days.Visible = False

' ---------------------------------------- picture initialization
Shape1.Height = Image1.Height
Shape1.Width = Image1.Width
ShowPicture

'---------------------------------------- ��־������������
'Call ShowMessage("Hello, Owner!", 1000)
'Call Delayms(2000)
Call ShowMessage("Hello, Owner!", 3000)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Dim n As Long
Dim Str As String

'---------------------------------------- ����������ӵ���֪ʶ�㣨cmdADD����Чʱ��Ч��
If cmdAdd.Enabled = True Then   '-----�����䡿���������ʱ�Ĳ��ȳ��򣬿��ܳ�Сbug
    cmdAdd_Click
End If

'---------------------------------------- save data

'-------------------- �������
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

'-------------------- ����֪ʶ�����ݿ�����
Open App.Path & "\Data\data_txt_group\data_v2__" & Data_Text_Idx & ".txt" For Output As #1
For n = 1 To DataNum
    Write #1, "__check__", Klg_Q(n), Klg_A(n), Klg_L(n), Klg_D(n)
Next
Close #1
Str = "-- ֪ʶ�����ݿ��ѱ��浽data_v2__" & Data_Text_Idx & ".txt"


'-------------------- ���汾��ѧϰ��������
If StudyNumReal > 0 Then    '���ʵ��ѧϰ��������0�����κ�һ������ѧϰ��ť��������>0
    Open App.Path & "\Data\studying.txt" For Output As #3
    For n = 1 To StudyNumReal
        Write #3, Study_Index(n), Study_Whether(n)
    Next
    'n = MsgBox("����ѧϰ�����ѱ�����studying.txt", 0 + 64, "Goodbye")
    Close #3
    Str = Str & Chr(13) & Chr(10) & "-- ����Ŀǰѧϰ�����ѱ�����studying.txt"
End If

'-------------------- ����֪ʶ�㱸�ݣ�2024-04�������һ��data�ļ��洢��backup��̭��
'If Flag_IfAdd = 1 Then    '������֪ʶ��ʱ����һ�Σ��������ٸ��´���
'    Open App.Path & "\Data\data_backup.txt" For Output As #4
'    For n = 1 To DataNum
'        Write #4, "__check__", Klg_Q(n), Klg_A(n), Klg_L(n), Klg_D(n)
'    Next
'    Close #4
'    Str = Str & Chr(13) & Chr(10) & "-- data_backup.txt�Ѹ������飡��"
'End If

'��ʾ��-----------------------------------------------------------
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
lblLV.Caption = ""            '���صȼ���ǩ��
txt_Days.Visible = False       '���������ı���
If CmdRevisePlan.Caption = "���ؽ��ϰ�ƻ�" Then CmdRevisePlan_Click  '�����б�

OrdinalNum = OrdinalNum + 1
If OrdinalNum > StudyNumReal Then 'study finished
    txtQuestion.Text = "����ѧϰ����ɣ��ڴ����´�ѧϰ�����κ��£����������Ҫ�ģ�"
    txtAnswer.Text = ""
    Call Cmd_EnableManage(FORM_STUDY_FINISH)
    ShowDataState      '��ʾ��֪ʶ������ѧ֪ʶ��
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
    '�����������ա���ָ�����ܴﵽ12���֪ʶ��
Next
lblAll.Caption = "֪ʶ������" & DataNum
lbllearned.Caption = "��ѧ֪ʶ��" & LearnedNum
lblacquired.Caption = "�������գ�" & AcquiredNum
End Sub

'function: ͨ�ù��̡���ʾһЩ��ʾ��Ϣ����������ʱ��Timer1
Private Sub ShowMessage(Str As String, ms As Integer)
txtMessage.Text = Str
txtMessage.Visible = True
Timer1.Interval = ms
Timer1.Enabled = True

End Sub


Private Sub Label2_Click()

End Sub

'function: Timer1��ʱ�����¼�����Ϣ����ʧ
Private Sub Timer1_Timer()
'Static t As Byte
Timer1.Enabled = False
txtMessage.Visible = False    '��Ϣ���Ϊ���ɼ�

End Sub

' ���ܣ���ȡĳ·���������ļ���
' ������list(): �ַ������飬���ڴ�ŷ��ص��ļ���
'       sPath: ·��
'       Filter����׺�޶�
' �����list()
Sub GetFileList(list() As String, ByVal sPath As String, ByVal Filter As String)
    '���ǻ�ȡָ��du�ļ�����ָ����׺������zhi�����ƵĹ��̣�װ������picname()��

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

'function: �Ҳ�ͼƬ�����ѡ����̡�
'notes: this Sub must be after the "key management" Sub
Private Sub ShowPicture()

Dim rand_idx As Integer
Dim pic_list() As String

Randomize    'init random seed
Call GetFileList(pic_list, App.Path & "\pictures\", "*.jpg,*.png")  '��ȡ·��������ͼƬ

' ------------------------------ ������ʾ��һ��ͼƬ������¼
If cmd1Next.Enabled = True Then  '���ǵ�����ν���ѧϰ��ͼƬ����
    Do
        rand_idx = Int(Rnd * (UBound(pic_list) - 1 + 1)) + 1    '����1����ΪͼƬ��01��ʼ����
    Loop Until rand_idx <> Pic_LastIndex
    Pic_LastIndex = rand_idx
End If

Image1.Picture = LoadPicture(App.Path & "\Pictures\" & pic_list(Pic_LastIndex)) 'show picture

End Sub


'function: �������״̬�İ�����ʹ�ܡ���ʾ����ͳһ����
Private Sub Cmd_EnableManage(which As Byte)
Dim i As Integer

'------------------------------ �ع鰴������״̬
cmd1This.Enabled = False
cmd1Next.Enabled = False

cmdShowAnswer.Enabled = False
cmdKnow.Enabled = False
cmdKnow.Caption = "֪  ��"
cmdDontKnow.Enabled = False
cmdDontKnow.Caption = "ģ  ��"
For i = 0 To 3
    cmd_ManualLvlChg.Item(i).Enabled = False '�ؼ��������Բ���ʵ��
Next

cmdChg.Enabled = False
cmdAdd.Enabled = False
cmdLrnNum.Enabled = False   '���ܲ�����ť��Ԥ������ѧϰ������
cmdOneMore.Enabled = False  '���ܲ�����ǿ����ѧ����ť
chk_Priority_idx.Enabled = False


'------------------------------ ���ݲ�ͬ�Ľ����������״̬
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
        cmdKnow.Caption = "֪ ����" & ProLvl_ifKnow & ")"
        cmdDontKnow.Enabled = True
        cmdDontKnow.Caption = "ģ ����" & ProLvl_ifDontKnow & ")"
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
Savetime = timeGetTime '���¿�ʼʱ��ʱ��
While timeGetTime < Savetime + ms 'ѭ���ȴ�
    DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼���
Wend
End Sub




