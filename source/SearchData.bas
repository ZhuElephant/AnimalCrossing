Attribute VB_Name = "SearchData"
Option Explicit

'COPY RIGHT:SW-2829-4681-6006
'Email:z1s1w1@gmail.com
'VERSION:   V0.02
'Date:      2020/4/1
'V0.01 StevenZhu The first version
'V0.02 StevenZhu Add mandarin language

Private Date_inf As Byte
Private Position_inf As String
Private Hour_inf As String
Private Insect_lst() As String
Private Fish_lst() As String
Private Insect_Row As Integer
Private Insect_Row_St As Integer
Private Fish_Row As Integer
Private Fish_Row_St As Integer

Private Sub Judge_Hour()
    Dim ActvHour As String
    Dim A
    Dim i As Byte
    ActvHour = "4時～8時"
    Hour_inf = "6"
    
    Dim reg As Object
    Dim reg1 As Object
    Dim reg2 As Object
    Set reg = CreateObject("VbScript.regexp")
    Set reg1 = CreateObject("VbScript.regexp")
    Set reg2 = CreateObject("VbScript.regexp")
    
    With reg
    .Global = True
    .IgnoreCase = False
    '.Pattern = "\d+時(～\d+時)?|\d+時(～\d+時)?|\d+時"
    .Pattern = "\d+時～\d+時"
    End With

    If (reg.Execute(ActvHour)) Then
        Split (ActvHour)
    
    End If
    
    
    
    
    
    Set A = reg.Execute(ActvHour)
    'Debug.Print A(0), A(1), A.Count
    
    For i = 0 To A.Count - 1
        Debug.Print Split(A(i), "時", 1)(0)
    Next
    
End Sub

Private Function JudgeMonRange(A As String) As Boolean
    Dim Month_Range As String
    Dim Month_St As String
    Dim Month_Ed As String
    Dim MonthSt As Byte
    Dim MonthEd As Byte
    Dim result As Boolean
    
    Month_Range = Split(A, "：", 2)(1)
    Month_St = Split(Month_Range, "～", 2)(0)
    Month_Ed = Split(Month_Range, "～", 2)(1)
    Month_St = Split(Month_St, "月", 2)(0)
    Month_Ed = Split(Month_Ed, "月", 2)(0)
    MonthSt = Val(Month_St)
    MonthEd = Val(Month_Ed)
    If MonthSt > MonthEd Then
        If (Date_inf + 12) >= MonthSt And (Date_inf + 12) <= (MonthEd + 12) Then
            result = True
        End If
    Else
        If Date_inf >= MonthSt And Date_inf <= MonthEd Then
            result = True
        End If
    End If
    
    JudgeMonRange = result
End Function
Private Function Judge_Pos(A As String) As Boolean
    Dim B As String
    Dim reg As Object
    Dim reg1 As Object
    Dim reg2 As Object
    Dim result As Boolean
    
    'X月～X月
    Set reg = CreateObject("VbScript.regexp")
    With reg
    .Global = True
    .IgnoreCase = False
    .Pattern = "\d+月～\d+月"
    End With

    If reg.Test(A) Then
        result = JudgeMonRange(A)
    End If
    
     '1年中
    Set reg1 = CreateObject("VbScript.regexp")
    With reg1
    .Global = True
    .IgnoreCase = False
    .Pattern = "1年中"
    End With
    
    If reg1.Test(A) Then
        result = True
    End If
    
    '3月～6月、9月～10月
    Set reg2 = CreateObject("VbScript.regexp")
    With reg2
    .Global = True
    .IgnoreCase = False
    .Pattern = "\d+月～\d+月、\d+月～\d+月"
    End With
    
    If reg2.Test(A) Then
        B = Split(A, "、", 2)(0)
        result = JudgeMonRange(B)
        If result = False Then
            B = "南半球：" + Split(A, "、", 2)(1)
            result = JudgeMonRange(B)
        End If
    End If
    
    Judge_Pos = result
End Function

Private Sub Fish_search()
    Dim i As Byte
    Dim j As Byte
    Dim x As Integer
    Dim ActvHour As String
    Dim ActvPos As String
    Dim Pos_OK As Boolean
    Dim Pos_time As String
    
    For j = 0 To UBound(Fish_lst)
        For i = 2 To 162 Step 2
            If Sheet3.Cells(i, "A") = Fish_lst(j) Then
            If Position_inf = "North" Then
                Pos_time = Sheet3.Cells(i, "D")
                Pos_OK = Judge_Pos(Pos_time)
            Else
                Pos_time = Sheet3.Cells(i + 1, "D")
                Pos_OK = Judge_Pos(Pos_time)
            End If

                'ActvHour = Sheet3.Cells(i, "E")
                
                If Pos_OK Then
                    x = Fish_Row_St + j
                    Sheet1.Cells(x, "B") = Sheet3.Cells(i, "B")
                    Sheet1.Cells(x, "C") = Pos_time
                    Sheet1.Cells(x, "D") = Sheet3.Cells(i, "E")
                    Sheet1.Cells(x, "E") = Sheet3.Cells(i, "F")
                    Sheet1.Cells(x, "F") = Sheet3.Cells(i, "G")
                End If
            End If
        Next
    
    Next
End Sub

Private Sub Insect_search()
    Dim i As Byte
    Dim j As Byte
    Dim x As Integer
    Dim ActvHour As String
    Dim ActvPos As String
    Dim Pos_OK As Boolean
    Dim Pos_time As String
    
    For j = 0 To UBound(Insect_lst)
        For i = 2 To 162 Step 2
            If Sheet3.Cells(i, "J") = Insect_lst(j) Then
            If Position_inf = "North" Then
                Pos_time = Sheet3.Cells(i, "M")
                Pos_OK = Judge_Pos(Pos_time)
            Else
                Pos_time = Sheet3.Cells(i + 1, "M")
                Pos_OK = Judge_Pos(Pos_time)
            End If
                If Pos_OK Then
                x = Insect_Row_St + j
                Sheet1.Cells(x, "B") = Sheet3.Cells(i, "K")
                Sheet1.Cells(x, "C") = Pos_time
                Sheet1.Cells(x, "D") = Sheet3.Cells(i, "N")
                Sheet1.Cells(x, "E") = Sheet3.Cells(i, "O")
                Sheet1.Cells(x, "F") = Sheet3.Cells(i, "P")
                End If
            End If
        Next
    
    Next
End Sub

Private Sub Sheet_Clear()
    Dim A As String

    If Insect_Row > Insect_Row_St Then
        A = "B" + Trim(Str(Insect_Row_St)) + ":F" + Trim(Str(Insect_Row))
        Sheet1.Range(A).ClearContents
    End If
    If Fish_Row > Fish_Row_St Then
        A = "B" + Trim(Str(Fish_Row_St)) + ":F" + Trim(Str(Fish_Row))
        Sheet1.Range(A).ClearContents
    End If
End Sub

Private Sub Get_ListInf()
    Dim x As Byte
    Dim y As Byte
    
    Insect_Row = 1
    Fish_Row = 1
    
    Do While Sheet1.Cells(Insect_Row, "A") <> "Insects you don't have"
        Insect_Row = Insect_Row + 1
    Loop

    Do While Sheet1.Cells(Fish_Row, "A") <> "Fishes you don't have"
        Fish_Row = Fish_Row + 1
    Loop
    
    Insect_Row_St = Insect_Row + 1
    Fish_Row_St = Fish_Row + 1
    Insect_Row = Insect_Row_St
    Fish_Row = Fish_Row_St
    
    Do While Sheet1.Cells(Insect_Row, "A") <> ""
        ReDim Preserve Insect_lst(x)
        Insect_lst(x) = Sheet1.Cells(Insect_Row, "A")
        Insect_Row = Insect_Row + 1
        x = x + 1
    Loop
        
    Do While Sheet1.Cells(Fish_Row, "A") <> ""
        ReDim Preserve Fish_lst(y)
        Fish_lst(y) = Sheet1.Cells(Fish_Row, "A")
        Fish_Row = Fish_Row + 1
        y = y + 1
    Loop
    Insect_Row = Insect_Row - 1
    Fish_Row = Fish_Row - 1
End Sub

Private Sub Get_BasicData()
    Position_inf = Sheet1.Range("A2")
    Date_inf = Sheet1.Range("B2")
    Hour_inf = Sheet1.Range("C2")
End Sub

Private Function ArrJudge(arr) As Boolean
    Dim i&
    On Error Resume Next
    i = UBound(arr)
    If Err = 0 Then ArrJudge = True
End Function

Private Sub COPY_RIGHT()
    Sheet1.Range("G1") = "COPY RIGHT:SW-2829-4681-6006"
    Sheet1.Range("G2") = "Email:z1s1w1@gmail.com"
    Sheet1.Range("G3") = "VERSION:   V0.02"
    Sheet1.Range("G4") = "Date:      2020/4/1"
End Sub


Private Sub Start_Search()
    Call Get_BasicData
    Call Get_ListInf
    Call Sheet_Clear
    If ArrJudge(Insect_lst) Then
        Call Insect_search
    End If
    If ArrJudge(Fish_lst) Then
        Call Fish_search
    End If
    Call COPY_RIGHT
End Sub

