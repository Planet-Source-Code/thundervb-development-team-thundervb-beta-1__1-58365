Attribute VB_Name = "modIsml"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
'Revision history:
'8/10/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'
'Many things have changed here ..
'just too bored to log anery small fix

' isml
Public Type Isml_list
    name As String
    items() As String
    count As Long
End Type

Public Enum Isml_kw_types
    kw_PopUpList = 1
    kw_Tipform = 2
    kw_both = 3
End Enum

Public Enum isml_match_types
    match_none
    match_String
    match_List
    match_Container
    match_Number
    match_UserVar
    match_UserLabel
    match_UserString
End Enum

Public Enum isml_match_mods
    match_none = 0
    match_nocase = 1
    match_any = 2
    match_Prev_Next = 4
    match_Next_Prev = 8
End Enum

Public Type isml_match_base
    mType As isml_match_types
    mMod As isml_match_mods
    mString As String
    mName As String
End Type

Public Type isml_match
    mType As isml_match_types
    mMod As isml_match_mods
    mString As String
    mName As String
    mList() As isml_match_base
    count As Long
End Type

Public Type Isml_kw
    Matches() As isml_match
    count As Long
    Isml_kw_Type As Isml_kw_types
End Type

Public Type Isml_File
    Lists() As Isml_list
    ListCount As Long
    kw() As Isml_kw
    kw_count As Long
End Type

Dim sLines() As String, numLines As Long

Public Function LoadIsmlFile(file As String) As Isml_File
Dim tl As String, cf As Long
    cf = FreeFile

    Open file For Binary As cf
        tl = Space(LOF(cf))
        Get cf, , tl
    Close cf
    
    sLines = Split(tl, vbNewLine)
    numLines = UBound(sLines)
    LoadIsmlFile = ParseIsmlFile()
    
End Function

Public Function ParseIsmlFile() As Isml_File
Dim i As Long, temp As Isml_File

    For i = 0 To numLines
        sLines(i) = Trim$(sLines(i))
        If Len(sLines(i)) = 0 Then GoTo NextOne
        If Mid$(sLines(i), 1, 1) = ";" Then GoTo NextOne
        
        If Mid$(sLines(i), 1, 2) = "%#" Then
            AddListToFile temp, GetList(i, temp)
        ElseIf Mid$(sLines(i), 1, 1) = "$" Then
            AddkwToFile temp, Getkw(i, temp)
        Else
            ErrorBox "Error in line # " & i & vbNewLine & _
                   "Unknown meaning " & Mid$(sLines(i), 1, 1) & vbNewLine & _
                   "Line string : " & sLines(i), "modIsml", "ParseIsmlFile"
            Exit Function
        End If
NextOne: 'skip to next line
    Next i
    
    ParseIsmlFile = temp
    
End Function

Public Function GetList(ByRef fromLine As Long, fromFile As Isml_File) As Isml_list
Dim temp As Isml_list
    
    temp.name = Mid$(sLines(fromLine), 2)
    fromLine = fromLine + 1
    
    With temp
        Do While fromLine <= numLines
            sLines(fromLine) = Trim$(sLines(fromLine))
            If Mid$(sLines(fromLine), 1, 1) <> "?" Then Exit Do
            If Mid$(sLines(fromLine), 2, 1) = "#" Then
                AddListToList temp, GetListFromName(Mid$(sLines(fromLine), 2), fromFile, fromLine)
            Else
                AddItemToList temp, Mid$(sLines(fromLine), 2)
            End If
        fromLine = fromLine + 1
        Loop
    End With
    GetList = temp
    fromLine = fromLine - 1
    
End Function

Public Sub AddListToFile(ToFile As Isml_File, list As Isml_list)

        
    With ToFile
        ReDim Preserve .Lists(.ListCount)
        .Lists(.ListCount) = list
        .ListCount = .ListCount + 1
    End With
    

End Sub

Public Sub AddkwToFile(ToFile As Isml_File, kw As Isml_kw)

        
    With ToFile
        ReDim Preserve .kw(.kw_count)
        .kw(.kw_count) = kw
        .kw_count = .kw_count + 1
    End With
    

End Sub

Public Sub AddListToList(ToList As Isml_list, addList As Isml_list)
Dim i As Long
    
    With addList
        For i = 0 To .count - 1
            AddItemToList ToList, .items(i)
        Next i
    End With
    
End Sub

Public Sub AddItemToList(ToList As Isml_list, item As String)
        
    With ToList
        ReDim Preserve .items(.count)
        .items(.count) = item
        .count = .count + 1
    End With
    
End Sub

Public Function GetListFromName(name As String, fromFile As Isml_File, ByVal curline As Long) As Isml_list
Dim i As Long

    With fromFile
        For i = 0 To .ListCount - 1
            If .Lists(i).name = name Then
                GetListFromName = .Lists(i)
                Exit Function
            End If
        Next i
    End With
    
    If name = "##" Then
        'GetListFromName = "##"
        Exit Function
    ElseIf name = "#$" Then
        'GetListFromName = "##"
        Exit Function
    End If
    
    ErrorBox "Error , List " & name & vbNewLine & _
           " is non defined or used in code before defined" & vbNewLine & _
           "On line " & curline, "modIsml", "GetListFromName"
           
End Function

Public Function ListToString(lst As Isml_list) As String
Dim i As Long, temp As String
    With lst
        
        For i = 0 To .count - 1
            temp = temp & .items(i)
            If .count > 1 And i < (.count - 1) Then
                temp = temp & "|"
            End If
        Next i
        
    End With
    ListToString = temp
End Function

Public Function Getkw(ByRef fromLine As Long, ByRef fromFile As Isml_File) As Isml_kw
Dim temp As Isml_kw, tline As String

    tline = sLines(fromLine)
    'fromLine = fromLine + 1
    
    With temp
        Select Case Mid$(tline, 2, 1)
            Case "!"
                .Isml_kw_Type = Isml_kw_types.kw_PopUpList
            Case "~"
                .Isml_kw_Type = kw_Tipform
            Case "*"
                .Isml_kw_Type = kw_both
            Case Else
                ErrorBox "Error , kw parsing line #" & fromLine - 1 & vbNewLine & _
                       Mid$(tline, 2, 1) & " is not a known type of kw" & vbNewLine & _
                       "On line " & tline, "modIsml", "Getkw"
                fromLine = ""
            Exit Function
        End Select
        tline = Trim$(Replace(Replace(Replace(Mid$(tline, 3), ",", " , "), "]", " ] "), "[", " [ "))
        Do While Len(tline)
            AddMachTokw temp, ParseMatch(fromFile, tline, fromLine - 1)
        Loop
    End With
    
    Getkw = temp
 
End Function

Sub AddMachTokw(tokw As Isml_kw, match As isml_match)

    With tokw
        ReDim Preserve .Matches(.count)
        .Matches(.count) = match
        .count = .count + 1
    End With
   
End Sub

Public Function ParseMatch(fromFile As Isml_File, sLine As String, nLine As Long) As isml_match
Dim temp As isml_match
    
    With temp
        
        Select Case Mid$(sLine, 1, 1)
            Case "@"
                .mType = match_String
                If Mid$(sLine, 2, Len("<nocase>")) = "<nocase>" Then
                    .mMod = .mMod Or match_nocase
                    Mid$(sLine, 1, Len("<nocase>") + 1) = Space(Len("<nocase>") + 1)
                    sLine = "@" & Trim$(sLine)
                End If
                '@<any>
                If Mid$(sLine, 2, Len("<any>")) = "<any>" Then
                    .mMod = .mMod Or match_any
                    Mid$(sLine, 1, Len("<any>") + 1) = Space(Len("<any>") + 1)
                    sLine = "@" & Trim$(sLine)
                End If
                '<prv><nxt>" ,"
                If Mid$(sLine, 2, Len("<prv><nxt>")) = "<prv><nxt>" Then
                    .mMod = .mMod Or match_Prev_Next
                    Mid$(sLine, 1, Len("<prv><nxt>") + 1) = Space(Len("<prv><nxt>") + 1)
                    sLine = "@" & Trim$(sLine)
                End If
                '<nxt><prv>" ,"
                If Mid$(sLine, 2, Len("<nxt><prv>")) = "<nxt><prv>" Then
                    .mMod = .mMod Or match_Next_Prev
                    Mid$(sLine, 1, Len("<nxt><prv>") + 1) = Space(Len("<nxt><prv>") + 1)
                    sLine = "@" & Trim$(sLine)
                End If
                
                .mString = Replace$(Mid$(sLine, 3, InStr(3, sLine, """") - 3), " , ", ",")
                Mid$(sLine, 1) = "   " + Space$(InStr(3, sLine, """") - 3)
                sLine = Trim$(sLine)
            Case "[" 'clild match list..
                .mType = match_Container
                Mid$(sLine, 1, 1) = " "
                sLine = Trim(sLine)
                Do
                    AddMatchToMatchListed temp, ParseMatch(fromFile, sLine, nLine)
                    If Mid$(sLine, 1, 1) = "," Then Mid$(sLine, 1, 1) = " ": sLine = Trim(sLine)
                Loop While Mid$(sLine, 1, 1) <> "]"
                Mid$(sLine, 1, 1) = " ": sLine = Trim(sLine)
            Case "#"
                Select Case Mid$(sLine, 2, 1)
                    ' ## is a number ,
                    Case "#"
                        .mType = match_Number
                        Mid$(sLine, 1, 2) = "  "
                        sLine = Trim$(sLine)
                    ' #$ is a var created by the user and
                    Case "$"
                        .mType = match_UserVar
                        Mid$(sLine, 1, 2) = "  "
                        sLine = Trim$(sLine)
                    ' #& is a label made by the user
                    Case "&"
                        .mType = match_UserLabel
                        Mid$(sLine, 1, 2) = "  "
                        sLine = Trim$(sLine)
                    ' #@ is a string made by the user
                    Case "@"
                        .mType = match_UserString
                        Mid$(sLine, 1, 2) = "  "
                        sLine = Trim$(sLine)
                    ' #name is a user defined list
                    Case Else
                        .mType = match_List
                        
                        If InStr(1, sLine, " ") - 1 > 0 Then
                            .mName = Left$(sLine, InStr(1, sLine, " ") - 1)
                        Else
                            .mName = sLine
                        End If
                        .mString = ListToString(GetListFromName(.mName, fromFile, nLine))
                        
                        If InStr(1, sLine, " ") - 1 > 0 Then
                            Mid$(sLine, 1, InStr(1, sLine, " ") - 1) = Space(InStr(1, sLine, " ") - 1)
                            sLine = Trim$(sLine)
                        Else
                            sLine = ""
                        End If
        
                        'MsgBox "Error , match parsing line #" & nLine & vbNewLine & _
                        '       Mid$(sLine, 2, 1) & " is not a known type of predefined type (##,#$,#&,#@)" & vbNewLine & _
                        '       "On line " & sLine
                End Select
                
            Case Else
                ErrorBox "Error , match parsing line #" & nLine & vbNewLine & _
                   Mid$(sLine, 1, 1) & " is not a known type of match" & vbNewLine & _
                   "On line " & sLine, "modIsml", "ParseMatch"
                   sLine = ""
                Exit Function
        End Select
    
    End With
    
    ParseMatch = temp
End Function

Public Sub AddMatchToMatchListed(ToMatch As isml_match, addV As isml_match)

    With ToMatch
        ReDim Preserve .mList(.count)
        .mList(.count).mMod = addV.mMod
        .mList(.count).mString = addV.mString
        .mList(.count).mType = addV.mType
        .mList(.count).mName = addV.mName
        .count = .count + 1
    End With
   
End Sub

Public Function kwListToString(kwlist As Isml_File, Optional ByVal Fromlevel As Long = 0, Optional ByVal numlevels As Long = -1, Optional Filter As Isml_kw_types = kw_both) As String
Dim i As Long, temp As String, temp2 As String
    With kwlist
        For i = 0 To .kw_count - 1
            temp2 = kwToString(.kw(i), Fromlevel, numlevels, Filter)
            'If Filter And kw_PopUpList Then
                'temp2 = Replace$(temp2, "|", " ")
            'End If
            temp = temp & temp2
            
            If i < (.kw_count - 1) And Len(temp2) > 0 Then
            
                temp = temp & "|"
            
            End If
            
        Next i
    End With
    kwListToString = temp
End Function


Public Function kwToString(kw As Isml_kw, Optional ByVal Fromlevel As Long = 0, Optional ByVal numlevels As Long = -1, Optional Filter As Isml_kw_types) As String
Dim i As Long, temp As String

    With kw
        If Fromlevel > (.count - 1) Then Fromlevel = .count - 1
        If Fromlevel > (.count - 1) Then Fromlevel = .count - 1
        If numlevels = -1 Then
            numlevels = (.count - Fromlevel) - 1
        ElseIf (Fromlevel + numlevels) > .count Then
            numlevels = (.count - Fromlevel) - 1
        End If
        
        If .Isml_kw_Type = kw_PopUpList Then
            If Filter = kw_Tipform Then Exit Function
        ElseIf .Isml_kw_Type = kw_Tipform Then
            If Filter = kw_PopUpList Then Exit Function
        End If
        
        For i = Fromlevel To Fromlevel + numlevels
            temp = temp & kwmtoString(.Matches(i), Filter)
            If i < (Fromlevel + numlevels) Then
                temp = temp & "|"
            End If
        Next i
    End With
    kwToString = temp
End Function

Public Function kwmblTokwm(kwmbl As isml_match_base) As isml_match

    With kwmblTokwm
        .mType = kwmbl.mType
        .mString = kwmbl.mString
        .mName = kwmbl.mName
        .mMod = kwmbl.mMod
    End With

End Function

Public Function kwmtoString(kwm As isml_match, ByVal msel As Isml_kw_types) As String
Dim temp As String
    With kwm
            Select Case .mType
                Case isml_match_types.match_Container
                    Dim i As Long
                    If msel = kw_Tipform Then
                        temp = "["
                    End If
                    For i = 0 To .count - 1
                        temp = temp & kwmtoString(kwmblTokwm(.mList(i)), msel)
                        If i < (.count - 1) Then
                            If msel = kw_Tipform Then
                                temp = temp & ","
                            Else
                                temp = temp & "|"
                            End If
                        End If
                    Next i
                    If msel = kw_Tipform Then
                        temp = temp & "]"
                    End If
                Case isml_match_types.match_List
                    If msel = kw_Tipform Then
                        temp = Right$(.mName, Len(.mName) - 1) 'return the name
                    Else
                        temp = .mString 'return the list
                    End If
                Case isml_match_types.match_none
                    temp = ""
                Case isml_match_types.match_Number
                    temp = "#"
                Case isml_match_types.match_String
                    temp = Replace(.mString, "_", " ")
                    If .mMod And match_any Then
                         temp = Replace(Trim$(.mString), "_", " ")
                    End If
                Case isml_match_types.match_UserLabel
                    temp = "label"
                Case isml_match_types.match_UserString
                    temp = "string"
                Case isml_match_types.match_UserVar
                    temp = "var"
            End Select

    End With
    kwmtoString = temp
End Function

Public Function isMatch(str As String, mEn As isml_match, Optional ttip As Boolean = False, Optional ByRef bNc As Long = 0) As Boolean
    
    Select Case mEn.mType
        Case isml_match_types.match_Container
        Dim i As Long
             For i = 0 To mEn.count - 1
                 If isMatch(str, kwmblTokwm(mEn.mList(i)), ttip) Then
                    isMatch = True
                    Exit For
                 End If
             Next i
             
        Case isml_match_types.match_List
            If (InStr(1, str, mEn.mString) > 0) Or (ttip = True) Then
                isMatch = True
            End If
            
        Case isml_match_types.match_none
            If (Len(str) = 0) Or (ttip = True) Then
                isMatch = True
            End If
        Case isml_match_types.match_Number
            If IsNumeric(str) Or ttip = True Then
                isMatch = True
            End If
        Case isml_match_types.match_String
            If mEn.mMod And match_nocase Then 'case non sensitive
                If Trim$(str) = LCase$(Replace(mEn.mString, "_", " ")) Then
                    isMatch = True
                End If
            Else 'case sesnitive
                If Len(Trim$(str)) = 0 And Len(str) > 0 Then
                    str = "_"
                End If
                'match any char in str
                If mEn.mMod And match_any Then
                    Dim tr As Long
                    tr = InStr(1, mEn.mString, Replace(Trim$(str), "_", " "), vbTextCompare)
                    
                    If mEn.mMod And match_Next_Prev Then ' 1=next , 2=prev
                        If tr = 1 Then
                            bNc = 1
                        Else
                            bNc = -1
                        End If
                    Else
                        If tr = 1 Then ' 1=prev , 2=next
                            bNc = -1
                        Else
                            bNc = 1
                        End If
                    End If
                    
                    If tr > 0 Then
                        isMatch = True
                    End If
                Else
                    If Replace(Trim$(str), "_", " ") = Replace(mEn.mString, "_", " ") Then
                        isMatch = True
                    End If
                End If
            End If
        Case isml_match_types.match_UserLabel
            'to be made lat8r
            isMatch = True
        Case isml_match_types.match_UserString
            'to be made lat8r
            isMatch = True
        Case isml_match_types.match_UserVar
            'to be made lat8r
            isMatch = True
    End Select

End Function

