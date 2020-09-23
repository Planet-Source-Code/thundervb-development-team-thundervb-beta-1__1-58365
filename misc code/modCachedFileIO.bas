Attribute VB_Name = "modCachedFileIO"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'Made by Raziel(29/8/2004[dd/mm/yyyy]) .. ten days after string builder ;)
'Based on the StringBuilder
'Simple file io using a buffer to cache data....
'use it as you wish , gime a credit

Public Type file_b
    filenum As Long
    maxbuflen As Long
    buflen As Long
    buf As String_B
End Type

Sub PrintToFile(file As file_b, data As String)
    
    With file
        AppendString .buf, data & vbNewLine
        .buflen = .buflen + Len(data & vbNewLine)
        
        If .buflen > .maxbuflen Then
            Put #.filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With
    
End Sub

Sub AppendToFile(file As file_b, data As String)
    
    With file
        AppendString .buf, data
        .buflen = .buflen + Len(data)
        
        If .buflen > .maxbuflen Then
            Put .filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With
    
End Sub


Sub FlushFile(file As file_b)

    With file
       
        If .buflen > 0 Then
            Put .filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With

End Sub

Function OpenFile(filename As String, Optional buffersize As Long = 32768) As file_b
Dim temp As file_b
    
    temp.filenum = FreeFile
    temp.maxbuflen = buffersize
    Open filename For Binary As temp.filenum
    OpenFile = temp
    
End Function

Sub CloseFile(file As file_b)
Dim nullF As file_b
    
    FlushFile file
    Close file.filenum
    file = nullF
    
End Sub
