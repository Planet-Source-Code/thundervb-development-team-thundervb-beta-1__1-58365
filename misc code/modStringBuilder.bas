Attribute VB_Name = "modStringBuilder"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'Made by Raziel(19/8/2004[dd/mm/yyyy])
'A Simple but yet effective String Builder
'Giving a speed boost on big strings creation...
'Also , this code is used to create colored strings
'use it as you wish , gime  a credit

Public Type String_B
    str() As String
    str_index As Long
    str_bound As Long
End Type

Public Type Col_String_entry
    strlen As Long
    col As Long
End Type

Public Type Col_String
    str() As Col_String_entry
    str_index As Long
    str_bound As Long
End Type

'For normal stringz
Sub AppendString(toString As String_B, data As String)

    With toString
        If .str_index >= .str_bound Then
            If .str_bound = 0 Then .str_bound = 1
            ReDim Preserve .str(.str_bound * 2)
            .str_bound = UBound(.str)
        End If
        .str(.str_index) = data
        .str_index = .str_index + 1
    End With
    
End Sub

Sub FinaliseString(toString As String_B)

    With toString
        ReDim Preserve .str(.str_index - 1)
        .str_bound = UBound(.str)
    End With
    
End Sub

Function GetString(fromString As String_B) As String

    With fromString
        ReDim Preserve .str(.str_index - 1)
        GetString = Join$(.str, "")
        .str_bound = UBound(.str)
    End With
    
End Function


'for color code...
Sub AppendColString(toString As Col_String, data As Long, col As Long)
Dim temp As Col_String_entry

    With toString
        If .str_index >= .str_bound Then
            If .str_bound = 0 Then .str_bound = 1
            ReDim Preserve .str(.str_bound * 2)
            .str_bound = UBound(.str)
        End If
        temp.strlen = data
        temp.col = col
        .str(.str_index) = temp
        .str_index = .str_index + 1
    End With
    
End Sub

Sub FinaliseColString(toString As Col_String)

    With toString
        ReDim Preserve .str(.str_index - 1)
    End With
    
End Sub
