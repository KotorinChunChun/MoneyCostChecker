Attribute VB_Name = "utlMSForms_Partial"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        UtlMSForms
Rem
Rem  @description   MSFormsのイケてないコントロールを、イイ感じに使うための関数群
Rem                 から抽出したもの
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
                                            ByVal hwnd As LongPtr, _
                                            ByVal hWndInsertAfter As Long, _
                                            ByVal x As Long, _
                                            ByVal y As Long, _
                                            ByVal cx As Long, _
                                            ByVal cy As Long, _
                                            ByVal wFlags As Long _
                                            ) As Long
#Else
    Private Declare Function SetWindowPos Lib "user32" ( _
                                            ByVal hwnd As LongPtr, _
                                            ByVal hWndInsertAfter As Long, _
                                            ByVal x As Long, _
                                            ByVal y As Long, _
                                            ByVal cx As Long, _
                                            ByVal cy As Long, _
                                            ByVal wFlags As Long _
                                            ) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                                            ByVal lpClassName As String, _
                                            ByVal lpWindowName As String _
                                            ) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                                            ByVal lpClassName As String, _
                                            ByVal lpWindowName As String _
                                            ) As Long
#End If
        
Const SWP_NOSIZE = &H1       'サイズ変更しない
Const SWP_NOMOVE = &H2       '位置変更しない
Const SWP_SHOWWINDOW = &H40  'ウィンドウを表示

Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

Rem フォームを常に最前面に表示
Rem
Rem  @param  fm          ユーザーフォームオブジェクト
Rem  @param  top_most    最前面表示するか否か
Rem
Public Sub UserForm_TopMost(fm As MSForms.UserForm, top_most As Boolean)
    Dim fmHWnd As LongPtr
    fmHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, fm.Caption)
    If fmHWnd = 0 Then Debug.Print Err.LastDllError: Err.Raise 9999, , "FindWindow Faild"
    
    If top_most Then
        Call SetWindowPos(fmHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(fmHWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End If
End Sub

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
Public Function ListBox_GetSelectedItemsDictionary(lb As MSForms.ListBox) As Object 'Dictionary
    Dim retVal As Object  'Dictionary
    Set retVal = CreateObject("Scripting.Dictionary")
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
            retVal.add i, rowItem
        End If
    Next
    
    Set ListBox_GetSelectedItemsDictionary = retVal
End Function

Rem リストボックスの選択アイテムの先頭列を配列で取得
Rem
Rem @param lb   リストボックスオブジェクト
Rem
Rem @return As Variant/Variant(0 to #)  選択中のアイテムの先頭列の配列
Rem                                     非選択時:要素0の配列
Rem
Rem @note
Rem     ※文字列認識なので重複アイテムは無条件に全て取得します。
Rem     ※重複を許容できない場合はIndexsの方を使用してください。
Rem
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
