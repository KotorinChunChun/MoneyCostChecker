VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMoneyCostChecker 
   Caption         =   "UserForm1"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "FormMoneyCostChecker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
Rem    会議等にかけているコストをカウントするツール
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/05/09 : 初回版
Rem
Rem @note
Rem    問題点：編集機能、保存機能、拡大バグ、計測時間が適当すぎ
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

Private Sub Update予定人件費()
    Dim 所要時間 As Date
    所要時間 = 終了日時 - 開始日時
    
    Dim 予定人件費 As Long
    予定人件費 = 秒単価 * (所要時間 * 24 * 60 * 60)
    
    Label_PlanPrice.Caption = _
        "　所要時間：" & Format(所要時間, "hh:mm") & _
        "　予定人件費：" & Format(予定人件費, "0,000") & "円"
End Sub

Private Sub TextBox_EndTime_Change(): Call Update予定人件費: End Sub
Private Sub TextBox_StartTime_Change(): Call Update予定人件費: End Sub

Private Sub CommandButton_Add_Click()
    Dim 役職: 役職 = InputBox("役職", "", "ほにゃららさん")
    If 役職 = "" Then Exit Sub
    
    Dim 単価: 単価 = InputBox("時給？", "", "1000")
    If 単価 = "" Then Exit Sub
    If Not IsNumStr(単価) Then Exit Sub
    
'    Dim 人数: 人数 = "0"
    Dim 人数: 人数 = InputBox("人数", "", "0")
    If 人数 = "" Then Exit Sub
    If Not IsNumStr(人数) Then Exit Sub
    
    Dim item: item = Array(役職, 単価, 人数)
    Call ListBox_AddItem(ListBox1, item)
End Sub

Private Sub ListItemCountUp(Add As Long)
    Dim colItems As Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        ListBox1.List(rowIndex, 2) = WorksheetFunction.max(0, colItems(rowIndex)(2) + Add)
    Next
    
    Call Update予定人件費
End Sub

Private Sub CommandButton_Minus_Click(): Call ListItemCountUp(-1): End Sub
Private Sub CommandButton_Plus_Click(): Call ListItemCountUp(1): End Sub

Private Sub SpinButton_MemberCount_Change()
    Call ListItemCountUp(SpinButton_MemberCount.Value - SPIN_DEFAULT_VALUE)
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
End Sub

Private Property Get 秒単価() As Double
    Dim ret As Double
    Dim data: data = ListBox1.List
    Dim i As Long
    For i = LBound(data) To UBound(data)
        ret = ret + (data(i, 1) / 60 / 60 * data(i, 2))
    Next
    秒単価 = ret
End Property

Private Sub ToggleButton_StartEnd_Click()
    Do
        If Not ToggleButton_StartEnd.Value Then Exit Do
        If IsClosing Then Exit Do
        
        Dim 経過時間 As Date
        経過時間 = Now() - 開始日時
'        経過時間 = CDate("1:00:00")

        Dim 残り時間 As Date
        残り時間 = 終了日時 - Now()
        
        Dim 人件費 As Long
        人件費 = 秒単価 * (経過時間 * 24 * 60 * 60)
        
        Label_メッセージ.Caption = "経過時間：" & Format(経過時間, "hh:mm:ss") & _
                                    "　毎秒:" & Format(秒単価, "0.00") & "円" & _
                                    "　会議終了まで残り:" & 残り時間
        
        Label_人件費.Caption = Format(人件費, "0,000") & "円"
        
        Application.Wait [Now() + "00:00:00.1"]
        DoEvents
    Loop
End Sub

Private Property Get 開始日時() As Date
    Static Last開始日時
    On Error Resume Next
    開始日時 = CDate(TextBox_StartDate.Text & " " & TextBox_StartTime.Text & ":00")
    On Error GoTo 0
    If 開始日時 = 0 Then
        開始日時 = Last開始日時
        TextBox_StartTime.BackColor = vbRed
    Else
        Last開始日時 = 開始日時
        TextBox_StartTime.BackColor = vbWhite
    End If
End Property
Private Property Let 開始日(dt As Date)
    TextBox_StartDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let 開始時間(dt As Date)
    TextBox_StartTime.Text = Format(dt, "hh:mm")
End Property

Private Property Get 終了日時() As Date
    Static Last終了日時
    On Error Resume Next
    TextBox_EndTime.Text = StrConv(TextBox_EndTime.Text, vbNarrow)
    終了日時 = CDate(TextBox_EndDate.Text & " " & TextBox_EndTime.Text & ":00")
    On Error GoTo 0
    If 終了日時 = 0 Then
        終了日時 = Last終了日時
        TextBox_EndTime.BackColor = vbRed
    Else
        Last終了日時 = 終了日時
        TextBox_EndTime.BackColor = vbWhite
    End If
End Property
Private Property Let 終了日(dt As Date)
    TextBox_EndDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let 終了時間(dt As Date)
    TextBox_EndTime.Text = Format(dt, "hh:mm")
End Property

Private Sub ToggleButton_最前面_Click()
    UserForm_TopMost Me, ToggleButton_最前面.Value
End Sub

Private Sub ToggleButton_設定表示_Click(): UserForm_画面サイズ設定: End Sub
Private Sub SpinButton_Zoom_Change(): Call UserForm_画面サイズ設定: End Sub
Private Property Get 画面倍率() As Double
    画面倍率 = SpinButton_Zoom.Value / SPIN_DEFAULT_VALUE
End Property
Private Sub UserForm_画面サイズ設定()
    Me.Width = WIDTH_DEFAULT * 画面倍率
    Me.Height = IIf(ToggleButton_設定表示.Value, HEIGHT_MAX * 画面倍率, HEIGHT_MIN * 画面倍率)
    Me.Zoom = 100 * 画面倍率
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "会議をダラダラ続けるのを止めさせよう！"
    Label_メッセージ.Caption = ""

    ToggleButton_最前面.Value = False
    ToggleButton_設定表示.Value = True
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;30;30"
    End With
    
    ListBox_AddItem ListBox1, Array("部長", "3000", "0")
    ListBox_AddItem ListBox1, Array("課長", "2500", "0")
    ListBox_AddItem ListBox1, Array("係長", "1500", "0")
    ListBox_AddItem ListBox1, Array("社員", "1200", "0")
    
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
    SpinButton_Zoom.Value = SPIN_DEFAULT_VALUE
    
    開始日 = Now()
    終了日 = Now()
    開始時間 = Now()
    終了時間 = DateAdd("h", 1, Now())
    
End Sub
