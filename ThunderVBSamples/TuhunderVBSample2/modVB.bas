Attribute VB_Name = "modVB"
Option Explicit

'Convertions for Unicode/Ascii
'VB string -> C string convertion
Public Function ConvToCString(str As String) As Byte()
Dim temp() As Byte
    
    temp = StrConv(str, vbFromUnicode)    'convert to text ascii
    ReDim Preserve temp(UBound(temp) + 1) 'add one more digit at the end
    temp(UBound(temp)) = 0                'for the terminator...
    ConvToCString = temp                  'Return the result
    
End Function

'C strring -> VB string convertion
Public Function ConvToVBString(strarr() As Byte) As String
Dim i As Long
    
    For i = 1 To UBound(strarr)
        If strarr(i) = 0 Then
            ReDim Preserve strarr(i - 1)
            Exit For
        End If
    Next i
    ConvToVBString = StrConv(strarr, vbUnicode)
    
End Function

