Attribute VB_Name = "WorksheetFunction"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        WorksheetFunction
Rem
Rem  @description   Excel�ȊO��VBA�ł�Excel.WorksheetFunction�Ɠ����֐����g����֐��Q
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

'�w�肵���z��̂����ő�̒l��Ԃ�
Public Function Max(ParamArray Nums() As Variant) As Variant
    Dim Num As Variant
    Max = Nums(0)
    For Each Num In Nums
        If Max < Num Then Max = Num
    Next
End Function
