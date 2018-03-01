'''
''' これらを利用するには、
''' =PERSONAL.XLSB!関数名(...)
''' とする。
'''

''' セル数式を文字列比較する
Function exactFx(a As Range, B As Range)
    If a.Formula = B.Formula Then
        exactFx = True
    Else
        exactFx = False
    End If
End Function

''' セル範囲の文字列を連結する
Function sumStr(aRange As Range, Optional delimter = "")
    Dim tmp As String
    Dim r As Range
    For Each r In aRange
        tmp = tmp & r.Value & delimter
    Next
    sumStr = Left(tmp, Len(tmp) - Len(delimter))End Function
End Function


