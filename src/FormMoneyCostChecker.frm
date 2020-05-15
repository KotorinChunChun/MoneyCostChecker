VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMoneyCostChecker 
   Caption         =   "UserForm1"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "FormMoneyCostChecker.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FormMoneyCostChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        FormMoneyCostChecker
Rem
Rem  @description   Time is money.
Rem                 ��c���ɂ����Ă���R�X�g���J�E���g����c�[��
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @note
Rem     ���_�F�ҏW�@�\�A�ۑ��@�\�A�g��o�O�A�v�����Ԃ��K������
Rem
Rem --------------------------------------------------------------------------------

Option Explicit

Const WIDTH_DEFAULT = 400
Const HEIGHT_MIN = 160
Const HEIGHT_MAX = 320

Const SPIN_DEFAULT_VALUE = 50

Private IsClosing As Boolean
Private TotalCost_ As Double

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub Update�\��l����()
    Dim ��c�\�莞�� As Date
    ��c�\�莞�� = �I������ - �J�n����
    ToggleButton_StartEnd.Enabled = (��c�\�莞�� > 0)
    
    Dim �\��l���� As Long
    �\��l���� = �b�P�� * (��c�\�莞�� * 24 * 60 * 60)
    
    Label_PlanPrice.Caption = _
        "�@��c�\�莞�ԁF" & Format(��c�\�莞��, "hh:mm") & _
        "�@�\��l����F" & Format(�\��l����, "0,000") & "�~"
End Sub

Private Sub TextBox_EndDate_Change(): Call Update�\��l����: End Sub
Private Sub TextBox_StartDate_Change(): Call Update�\��l����: End Sub
Private Sub TextBox_EndTime_Change(): Call Update�\��l����: End Sub
Private Sub TextBox_StartTime_Change(): Call Update�\��l����: End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ListItemEdit
End Sub

Private Function Get��E�P��(ByRef ��E, ByRef �P��) As Boolean
    ��E = InputBox("��E", "", ��E)
    If ��E = "" Then Exit Function
    
    �P�� = InputBox("�l����", "", �P��)
    If �P�� = "" Then Exit Function
    If Not IsNumStr(�P��) Then Exit Function
    
    Get��E�P�� = True
End Function

Private Sub ListItemEdit()
    Dim colItems As Object 'Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        Dim ��E: ��E = colItems(rowIndex)(0)
        Dim �P��: �P�� = colItems(rowIndex)(1)
        If Get��E�P��(��E, �P��) Then
            ListBox1.List(rowIndex, 0) = ��E
            ListBox1.List(rowIndex, 1) = �P��
        End If
    Next
End Sub

Private Sub CommandButton_Add_Click()
    Dim ��E: ��E = "�قɂ��炳��"
    Dim �P��: �P�� = "1000"
    Dim �l��: �l�� = "0"
    If Not Get��E�P��(��E, �P��) Then Exit Sub
    
    Dim item: item = Array(��E, �P��, �l��)
    Call ListBox_AddItem(ListBox1, item)
End Sub

Private Sub ListItemCountUp(add As Long)
    Dim colItems As Object 'Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        ListBox1.List(rowIndex, 2) = WorksheetFunction.Max(0, colItems(rowIndex)(2) + add)
    Next
    
    Call Update�\��l����
End Sub

Private Sub CommandButton_Minus_Click(): Call ListItemCountUp(-1): End Sub
Private Sub CommandButton_Plus_Click(): Call ListItemCountUp(1): End Sub

Private Sub SpinButton_MemberCount_Change()
    Call ListItemCountUp(SpinButton_MemberCount.Value - SPIN_DEFAULT_VALUE)
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
End Sub

Private Property Get TotalCost() As Double: TotalCost = TotalCost_: End Property
Private Property Let TotalCost(set_cost As Double)
    TotalCost_ = set_cost
    CommandButton_Reset.Visible = (TotalCost_ > 0)
    Label_�l����.Caption = Format(TotalCost_, "#,##0") & "�~"
End Property
Private Sub TotalCost_Add(add_cost As Double)
    TotalCost = TotalCost + add_cost
End Sub

Private Sub CommandButton_Reset_Click()
    TotalCost = 0
End Sub

Private Property Get �b�P��() As Double
    Dim ret As Double
    Dim data: data = ListBox1.List
    Dim i As Long
    For i = LBound(data) To UBound(data)
        ret = ret + (data(i, 1) / 60 / 60 * data(i, 2))
    Next
    �b�P�� = ret
End Property

Private Sub ToggleButton_StartEnd_Click()
    ConfigCtrlEnabled = False
    Dim ��c�\�莞�� As Date
    ��c�\�莞�� = �I������ - �J�n����
    
    Dim �O�i�F�����l As Long: �O�i�F�����l = Label_�l����.ForeColor
    Dim �w�i�F�����l As Long: �w�i�F�����l = Label_�l����.BackColor
    
    '�J�n�҂�
    Do
        If �J�n���� < Now() Then Exit Do
        
        Dim �J�n�܂Ŏc�� As Date
        �J�n�܂Ŏc�� = �J�n���� - Now()

        Label_���b�Z�[�W.Caption = "��c�J�n�܂Ŏc�� : " & �J�n�܂Ŏc��
        
        Call Sleep(100)
        DoEvents
    Loop
    
    '�J�n��
    Do
        If Not ToggleButton_StartEnd.Value Then Exit Do
        If IsClosing Then Exit Do
        
        Dim �o�ߎ��� As Date
        �o�ߎ��� = Now() - �J�n����
'        �o�ߎ��� = CDate("1:00:00")

        Dim �c�莞�� As Date
        �c�莞�� = �I������ - Now()
        
        Label_���b�Z�[�W.Caption = "�o�ߎ��� :  " & Format(�o�ߎ���, "hh:mm:ss") & _
                                    "�@���b : " & Format(�b�P��, "0.00") & "�~" & _
                                    "�@��c�I���܂Ŏc�� : " & �c�莞��
        
        '10������������F��؂�ւ����
'        Label_�l����.ForeColor = IIf(�c�莞�� < (1 / 24 / 60 * 10), vbRed, vbBlack)
        '�i�����ɉ����ĐF���t�F�[�h�C���������
        Dim �i���� As Long
        �i���� = CLng(CDbl(�o�ߎ���) / CDbl(��c�\�莞��) * 100)
        If �i���� <= 100 Then
            Label_�l����.ForeColor = GetFadeColor(�O�i�F�����l, vbRed, �i����)
        Else
            If Label_�l����.BackColor = �w�i�F�����l Then
                Label_�l����.ForeColor = vbWhite
                Label_�l����.BackColor = vbRed
            Else
                Label_�l����.ForeColor = vbRed
                Label_�l����.BackColor = �w�i�F�����l
            End If
        End If
        
        Call TotalCost_Add(�b�P�� / 5)
        Call Sleep(200)
        DoEvents
    Loop
    Label_�l����.ForeColor = �O�i�F�����l
    Label_�l����.BackColor = �w�i�F�����l
    ConfigCtrlEnabled = True
End Sub

Private Property Let ConfigCtrlEnabled(set_enabled As Boolean)
    
    CommandButton_Reset.Enabled = set_enabled

    TextBox_EndDate.Enabled = set_enabled
    TextBox_EndTime.Enabled = set_enabled

    TextBox_StartDate.Enabled = set_enabled
    TextBox_StartTime.Enabled = set_enabled
    
    ListBox1.Enabled = set_enabled
    CommandButton_Add.Enabled = set_enabled
    SpinButton_MemberCount.Enabled = set_enabled
    
End Property

Private Property Get �J�n����() As Date
    Static Last�J�n����
    On Error Resume Next
    �J�n���� = CDate(TextBox_StartDate.Text & " " & TextBox_StartTime.Text & ":00")
    On Error GoTo 0
    If �J�n���� = 0 Then
        �J�n���� = Last�J�n����
        TextBox_StartDate.BackColor = vbRed
        TextBox_StartTime.BackColor = vbRed
    Else
        Last�J�n���� = �J�n����
        TextBox_StartDate.BackColor = vbWhite
        TextBox_StartTime.BackColor = vbWhite
    End If
End Property
Private Property Let �J�n��(dt As Date)
    TextBox_StartDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let �J�n����(dt As Date)
    TextBox_StartTime.Text = Format(dt, "hh:mm")
End Property

Private Property Get �I������() As Date
    Static Last�I������
    On Error Resume Next
    TextBox_EndTime.Text = StrConv(TextBox_EndTime.Text, vbNarrow)
    �I������ = CDate(TextBox_EndDate.Text & " " & TextBox_EndTime.Text & ":00")
    On Error GoTo 0
    If �I������ = 0 Then
        �I������ = Last�I������
        TextBox_EndDate.BackColor = vbRed
        TextBox_EndTime.BackColor = vbRed
    Else
        Last�I������ = �I������
        TextBox_EndDate.BackColor = vbWhite
        TextBox_EndTime.BackColor = vbWhite
    End If
End Property
Private Property Let �I����(dt As Date)
    TextBox_EndDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let �I������(dt As Date)
    TextBox_EndTime.Text = Format(dt, "hh:mm")
End Property

Private Sub ToggleButton_�őO��_Click()
    UserForm_TopMost Me, ToggleButton_�őO��.Value
End Sub

Private Sub ToggleButton_�ݒ�\��_Click(): UserForm_��ʃT�C�Y�ݒ�: End Sub
Private Sub SpinButton_Zoom_Change(): Call UserForm_��ʃT�C�Y�ݒ�: End Sub
Private Property Get ��ʔ{��() As Double
    ��ʔ{�� = SpinButton_Zoom.Value / SPIN_DEFAULT_VALUE
End Property
Private Sub UserForm_��ʃT�C�Y�ݒ�()
    Me.Width = WIDTH_DEFAULT * ��ʔ{��
    Me.Height = IIf(ToggleButton_�ݒ�\��.Value, HEIGHT_MAX * ��ʔ{��, HEIGHT_MIN * ��ʔ{��)
    Me.Zoom = 100 * ��ʔ{��
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "��c���_���_��������̂��~�߂����悤�I"
    Label_���b�Z�[�W.Caption = ""

    CommandButton_Reset.Visible = False
    ToggleButton_�őO��.Value = False
    ToggleButton_�ݒ�\��.Value = True
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;30;30"
    End With
    
    ListBox_AddItem ListBox1, Array("����", "6000", "0")
    ListBox_AddItem ListBox1, Array("�ے�", "5000", "0")
    ListBox_AddItem ListBox1, Array("�W��", "4000", "0")
    ListBox_AddItem ListBox1, Array("�Ј�", "3000", "0")
    
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
    SpinButton_Zoom.Value = SPIN_DEFAULT_VALUE
    
    �J�n�� = Now()
    �I���� = DateAdd("h", 1, Now())
    �J�n���� = Now()
    �I������ = DateAdd("h", 1, Now())
    
#If DEF_EXCEL Then
    Excel.Application.Visible = False
#ElseIf DEF_WORD Then
    Word.Application.Visible = False
#End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    IsClosing = True
#If DEF_EXCEL Then
    Excel.Application.Visible = True
#ElseIf DEF_WORD Then
    Word.Application.Visible = True
#End If
End Sub


