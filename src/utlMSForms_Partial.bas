Attribute VB_Name = "utlMSForms_Partial"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        UtlMSForms
Rem
Rem  @description   MSForms�̃C�P�ĂȂ��R���g���[�����A�C�C�����Ɏg�����߂̊֐��Q
Rem                 ���璊�o��������
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
        
Const SWP_NOSIZE = &H1       '�T�C�Y�ύX���Ȃ�
Const SWP_NOMOVE = &H2       '�ʒu�ύX���Ȃ�
Const SWP_SHOWWINDOW = &H40  '�E�B���h�E��\��

Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

Rem �t�H�[������ɍőO�ʂɕ\��
Rem
Rem  @param  fm          ���[�U�[�t�H�[���I�u�W�F�N�g
Rem  @param  top_most    �őO�ʕ\�����邩�ۂ�
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

Rem ���X�g�{�b�N�X�ɃA�C�e����ǉ�����
Rem
Rem
Rem  @note
Rem    �W����AddItem���\�b�h�͔z��ɑΉ����Ă��Ȃ����ߕK�v
Rem    �n���ꂽ�z��̗v�f�����A�\���\�ȗ񐔂𒴂��Ă��Ă��؂�̂Ă���
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

Rem ���X�g�{�b�N�X�̑I���A�C�e�����f�B�N�V���i���Ŏ擾
Rem  �R���N�V�����ɂ͈ꎟ���z����i�[���A�s���̗�����i�[����B
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

Rem ���X�g�{�b�N�X�̑I���A�C�e���̐擪���z��Ŏ擾
Rem
Rem @param lb   ���X�g�{�b�N�X�I�u�W�F�N�g
Rem
Rem @return As Variant/Variant(0 to #)  �I�𒆂̃A�C�e���̐擪��̔z��
Rem                                     ��I����:�v�f0�̔z��
Rem
Rem @note
Rem     ��������F���Ȃ̂ŏd���A�C�e���͖������ɑS�Ď擾���܂��B
Rem     ���d�������e�ł��Ȃ��ꍇ��Indexs�̕����g�p���Ă��������B
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
