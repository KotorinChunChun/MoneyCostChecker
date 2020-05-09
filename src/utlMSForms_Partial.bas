Attribute VB_Name = "utlMSForms_Partial"
Rem --------------------------------------------------
Rem
Rem @module     UtlMSForms_Partial
Rem
Rem @description
Rem    MSFormsのイケてないコントロールを、イイ感じに使うための関数群
Rem　　　から抽出したもの
Rem
Rem --------------------------------------------------
Option Explicit

Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
        
Const SWP_NOSIZE = &H1       'サイズ変更しない
Const SWP_NOMOVE = &H2       '位置変更しない
Const SWP_SHOWWINDOW = &H40  'ウィンドウを表示

Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

'フォームを常に最前面に表示
Function UserForm_TopMost(F As MSForms.UserForm, TopMost As Boolean)
    Dim fmHWnd As LongPtr
    fmHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, F.Caption)
    If fmHWnd = 0 Then Err.Raise 9999, , "FindWindow Faild": Debug.Print Err.LastDllError
    
    If TopMost Then
        Call SetWindowPos(fmHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(fmHWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End If
End Function

'キャスト
Public Function ToCheckBox(Ctrl As MSForms.control) As MSForms.CheckBox: Set ToCheckBox = Ctrl: End Function
Public Function ToComboBox(Ctrl As MSForms.control) As MSForms.ComboBox: Set ToComboBox = Ctrl: End Function
Public Function ToCommandButton(Ctrl As MSForms.control) As MSForms.CommandButton: Set ToCommandButton = Ctrl: End Function
Public Function ToFrame(Ctrl As MSForms.control) As MSForms.frame: Set ToFrame = Ctrl: End Function
Public Function ToImage(Ctrl As MSForms.control) As MSForms.Image: Set ToImage = Ctrl: End Function
Public Function ToLabel(Ctrl As MSForms.control) As MSForms.label: Set ToLabel = Ctrl: End Function
Public Function ToListBox(Ctrl As MSForms.control) As MSForms.ListBox: Set ToListBox = Ctrl: End Function
Public Function ToMultiPage(Ctrl As MSForms.control) As MSForms.MultiPage: Set ToMultiPage = Ctrl: End Function
Public Function ToOptionButton(Ctrl As MSForms.control) As MSForms.OptionButton: Set ToOptionButton = Ctrl: End Function
Public Function ToSpinButton(Ctrl As MSForms.control) As MSForms.SpinButton: Set ToSpinButton = Ctrl: End Function
Public Function ToTabStrip(Ctrl As MSForms.control) As MSForms.TabStrip: Set ToTabStrip = Ctrl: End Function
Public Function ToTextBox(Ctrl As MSForms.control) As MSForms.TextBox: Set ToTextBox = Ctrl: End Function
Public Function ToToggleButton(Ctrl As MSForms.control) As MSForms.ToggleButton: Set ToToggleButton = Ctrl: End Function

Rem リストボックスにアイテムを追加する
Rem
Rem
Rem  @note
Rem    標準のAddItemメソッドは配列に対応していないため必要
Rem    渡された配列の要素数が、表示可能な列数を超えていても切り捨てられる
Rem
Public Function ListBox_AddItem(lb As MSForms.ListBox, insertRowData, Optional ByVal insertRowIndex As Long = -1) As Long
    If insertRowIndex = -1 Then
        insertRowIndex = lb.ListCount
    End If
    
    If Not IsArray(insertRowData) Then
        lb.addItem insertRowData, insertRowIndex
        ListBox_AddItem = insertRowIndex
        Exit Function
    End If
    
    lb.addItem "", insertRowIndex
    Dim columnIndex As Long, itemIndex As Long
    itemIndex = LBound(insertRowData)
    For columnIndex = 0 To lb.ColumnCount - 1
        lb.List(insertRowIndex, columnIndex) = insertRowData(itemIndex)
        itemIndex = itemIndex + 1
    Next
    ListBox_AddItem = insertRowIndex
End Function

Rem リストボックスの選択アイテムをディクショナリで取得
Rem  コレクションには一次元配列を格納し、行毎の列情報を格納する。
Rem
Rem  @param
Rem
Rem  @return As Dictionary(row)(Column Array)
Rem
Public Function ListBox_GetSelectedItemsDictionary(lb As MSForms.ListBox) As Dictionary
    Dim retVal As New Dictionary
    Set ListBox_GetSelectedItemsDictionary = retVal
    If lb.ListCount = 0 Then Exit Function
    
    Dim rowItem()
    Dim i As Long, j As Long
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            ReDim rowItem(0 To lb.ColumnCount - 1)
            For j = 0 To lb.ColumnCount - 1
                rowItem(j) = lb.List(i, j)
            Next
            retVal.Add i, rowItem
        End If
    Next
    
    Set ListBox_GetSelectedItemsDictionary = retVal
End Function

'SelectedIndexsも欲しい

'リストボックスの選択アイテムの先頭列を配列で取得
'※文字列認識なので重複アイテムは無条件に全て取得します。
'非選択時:要素0の配列
Public Function ListBox_GetSelectedItems(lb As MSForms.ListBox) As Variant
    ListBox_GetSelectedItems = VBA.Array()
    If lb.ListCount = 0 Then Exit Function
    
    Dim arr
    ReDim arr(0 To lb.ListCount - 1)
    Dim i As Long, nextIndex As Long
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            arr(nextIndex) = lb.List(i)
            nextIndex = nextIndex + 1
        End If
    Next
    
    Dim listData
    listData = lb.List
    
    If nextIndex = 0 Then ListBox_GetSelectedItems = VBA.Array(): Exit Function
    ReDim Preserve arr(0 To nextIndex - 1)
    
    ListBox_GetSelectedItems = arr
End Function

