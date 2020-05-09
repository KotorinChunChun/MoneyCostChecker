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

Rem --------------------------------------------------
Rem
Rem @module FormMoneyCostChecker
Rem
Rem @description
Rem    Time is money.
Rem    ��c���ɂ����Ă���R�X�g���J�E���g����c�[��
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/05/09 : �����
Rem
Rem @note
Rem    ���_�F�ҏW�@�\�A�ۑ��@�\�A�g��o�O�A�v�����Ԃ��K������
Rem
Rem --------------------------------------------------

Option Explicit

Const WIDTH_DEFAULT = 400
Const HEIGHT_MIN = 160
Const HEIGHT_MAX = 320

Const SPIN_DEFAULT_VALUE = 50

Private IsClosing As Boolean

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    IsClosing = True
End Sub

Private Sub Update�\��l����()
    Dim ���v���� As Date
    ���v���� = �I������ - �J�n����
    
    Dim �\��l���� As Long
    �\��l���� = �b�P�� * (���v���� * 24 * 60 * 60)
    
    Label_PlanPrice.Caption = _
        "�@���v���ԁF" & Format(���v����, "hh:mm") & _
        "�@�\��l����F" & Format(�\��l����, "0,000") & "�~"
End Sub

Private Sub TextBox_EndTime_Change(): Call Update�\��l����: End Sub
Private Sub TextBox_StartTime_Change(): Call Update�\��l����: End Sub

Private Sub CommandButton_Add_Click()
    Dim ��E: ��E = InputBox("��E", "", "�قɂ��炳��")
    If ��E = "" Then Exit Sub
    
    Dim �P��: �P�� = InputBox("�����H", "", "1000")
    If �P�� = "" Then Exit Sub
    If Not IsNumStr(�P��) Then Exit Sub
    
'    Dim �l��: �l�� = "0"
    Dim �l��: �l�� = InputBox("�l��", "", "0")
    If �l�� = "" Then Exit Sub
    If Not IsNumStr(�l��) Then Exit Sub
    
    Dim item: item = Array(��E, �P��, �l��)
    Call ListBox_AddItem(ListBox1, item)
End Sub

Private Sub ListItemCountUp(Add As Long)
    Dim colItems As Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        ListBox1.List(rowIndex, 2) = WorksheetFunction.max(0, colItems(rowIndex)(2) + Add)
    Next
    
    Call Update�\��l����
End Sub

Private Sub CommandButton_Minus_Click(): Call ListItemCountUp(-1): End Sub
Private Sub CommandButton_Plus_Click(): Call ListItemCountUp(1): End Sub

Private Sub SpinButton_MemberCount_Change()
    Call ListItemCountUp(SpinButton_MemberCount.Value - SPIN_DEFAULT_VALUE)
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
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
    Do
        If Not ToggleButton_StartEnd.Value Then Exit Do
        If IsClosing Then Exit Do
        
        Dim �o�ߎ��� As Date
        �o�ߎ��� = Now() - �J�n����
'        �o�ߎ��� = CDate("1:00:00")

        Dim �c�莞�� As Date
        �c�莞�� = �I������ - Now()
        
        Dim �l���� As Long
        �l���� = �b�P�� * (�o�ߎ��� * 24 * 60 * 60)
        
        Label_���b�Z�[�W.Caption = "�o�ߎ��ԁF" & Format(�o�ߎ���, "hh:mm:ss") & _
                                    "�@���b:" & Format(�b�P��, "0.00") & "�~" & _
                                    "�@��c�I���܂Ŏc��:" & �c�莞��
        
        Label_�l����.Caption = Format(�l����, "0,000") & "�~"
        
        Application.Wait [Now() + "00:00:00.1"]
        DoEvents
    Loop
End Sub

Private Property Get �J�n����() As Date
    Static Last�J�n����
    On Error Resume Next
    �J�n���� = CDate(TextBox_StartDate.Text & " " & TextBox_StartTime.Text & ":00")
    On Error GoTo 0
    If �J�n���� = 0 Then
        �J�n���� = Last�J�n����
        TextBox_StartTime.BackColor = vbRed
    Else
        Last�J�n���� = �J�n����
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
        TextBox_EndTime.BackColor = vbRed
    Else
        Last�I������ = �I������
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

    ToggleButton_�őO��.Value = False
    ToggleButton_�ݒ�\��.Value = True
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;30;30"
    End With
    
    ListBox_AddItem ListBox1, Array("����", "3000", "0")
    ListBox_AddItem ListBox1, Array("�ے�", "2500", "0")
    ListBox_AddItem ListBox1, Array("�W��", "1500", "0")
    ListBox_AddItem ListBox1, Array("�Ј�", "1200", "0")
    
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
    SpinButton_Zoom.Value = SPIN_DEFAULT_VALUE
    
    �J�n�� = Now()
    �I���� = Now()
    �J�n���� = Now()
    �I������ = DateAdd("h", 1, Now())
    
End Sub
