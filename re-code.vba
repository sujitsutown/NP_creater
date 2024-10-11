Sub Create()

    sjtStartRow = 1
    sjtStartCol = 1
    sjtEndRow = 4
    sjtEndCol = 4
    userNumber = 1
    
    '' 編集しないでください
    bushomei = "代表社員"

    lastRow = Worksheets("全体").Range("A1").End(xlDown).Row
    lastNumber = Worksheets("全体").Cells(lastRow, 1).Value

    ResetAllPageBreaks

    For i = 1 To lastNumber - 1
        If Worksheets("全体").Cells(i + 1, 6).Value = "社長" Or Worksheets("全体").Cells(i + 1, 6).Value = "副社長" Then
            '' 変数名をnameにすると誤作動が起きます
            namae = Worksheets("全体").Cells(i + 1, 5).Value
            dep = Worksheets("全体").Cells(i + 1, 6).Value
            about = Worksheets("全体").Cells(i + 1, 7).Value
            Call Bc_Create(namae, dep, about, sjtStartRow, sjtStartCol, sjtEndRow, sjtEndCol, bushomei, userNumber)
            sjtStartRow = sjtStartRow + 13
            sjtEndRow = sjtEndRow + 13
            userNumber = userNumber + 1
        End If
    Next i

End Sub

Public Sub Bc_Create(ByVal namae As String, ByVal dep As String, ByVal about As String, ByVal sjtStartRow As Integer, ByVal sjtStartCol As Integer, ByVal sjtEndRow As Integer, ByVal sjtEndCol As Integer, ByVal bushomei As String, ByVal userNumber As Integer)

    imgPath = ThisWorkbook.Path & "\校章.png"

    '' イベント名の記入
    Cells(sjtStartRow, sjtStartCol).Font.Color = RGB(250, 250, 250)
    Cells(sjtStartRow, sjtStartCol).HorizontalAlignment = xlCenter
    Cells(sjtStartRow, sjtStartCol).VerticalAlignment = xlCenter
    Cells(sjtStartRow, sjtStartCol).Interior.Color = RGB(32, 32, 32)
    Range(Cells(sjtStartRow, sjtStartCol), Cells(sjtEndRow, sjtEndCol)).Merge
    Cells(sjtStartRow, sjtStartCol).Value = "諏実タウン"
    Cells(sjtStartRow, sjtStartCol).Font.Size = 18

    '' 部署名の記入
    depStartrow = sjtStartRow + 4
    depStartCol = sjtStartCol
    depEndRow = sjtEndRow + 2
    depEndCol = sjtEndCol - 1
    
    Cells(depStartrow, depStartCol).HorizontalAlignment = xlCenter
    Cells(depStartrow, depStartCol).VerticalAlignment = xlCenter
    Cells(depStartrow, depStartCol).Font.Color = RGB(32, 32, 32)
    Range(Cells(depStartrow, depStartCol), Cells(depEndRow, depEndCol)).Merge
    If about = "" Then
        Cells(depStartrow, depStartCol).Value = dep
    Else
        Cells(depStartrow, depStartCol).Value = about
    End If
    Cells(depStartrow, depStartCol).Font.Size = 18
    
    '' 氏名の記入
    nameStartRow = sjtStartRow + 6
    nameStartCol = sjtStartCol
    nameEndRow = sjtEndRow + 5
    nameEndCol = sjtEndCol + 1
    
    Cells(nameStartRow, nameStartCol).HorizontalAlignment = xlCenter
    Cells(nameStartRow, nameStartCol).VerticalAlignment = xlCenter
    Cells(nameStartRow, nameStartCol).Font.Color = RGB(32, 32, 32)
    Range(Cells(nameStartRow, nameStartCol), Cells(nameEndRow, nameEndCol)).Merge
    Cells(nameStartRow, nameStartCol).Value = namae
    Cells(nameStartRow, nameStartCol).Font.Size = 26
    
    '' 休憩中の記入
    kyuStartRow = sjtStartRow
    kyuStartCol = sjtStartCol + 5
    kyuEndRow = sjtEndRow + 7
    kyuEndCol = sjtEndCol + 6
    
    Cells(kyuStartRow, kyuStartCol).HorizontalAlignment = xlCenter
    Cells(kyuStartRow, kyuStartCol).VerticalAlignment = xlCenter
    Cells(kyuStartRow, kyuStartCol).Font.Color = RGB(32, 32, 32)
    Range(Cells(kyuStartRow, kyuStartCol), Cells(kyuEndRow, kyuEndCol)).Merge
    Cells(kyuStartRow, kyuStartCol).Value = "休憩中"
    Cells(kyuStartRow, kyuStartCol).Font.name = "游ゴシック"
    Cells(kyuStartRow, kyuStartCol).Font.Size = 72

    '' 外枠の記入
    Range(Cells(kyuStartRow, kyuStartCol), Cells(kyuEndRow, kyuEndCol)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    Range(Cells(sjtStartRow, sjtStartCol), Cells(kyuEndRow, kyuEndCol)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium



    '' 校章の記入
    seStartRow = sjtStartRow
    seStartCol = sjtStartCol + 4
    seEndRow = sjtEndRow
    seEndCol = sjtEndCol + 1
    
    Cells(seStartRow, seStartCol).HorizontalAlignment = xlCenter
    Cells(seStartRow, seStartCol).VerticalAlignment = xlCenter
    Range(Cells(seStartRow, seStartCol), Cells(seEndRow, seEndCol)).Merge
    With Sheets(bushomei).Pictures.Insert(imgPath)
        .Left = Sheets(bushomei).Range(Cells(seStartRow, seStartCol), Cells(seEndRow, seEndCol)).Left
        .Top = Sheets(bushomei).Range(Cells(seStartRow, seStartCol), Cells(seEndRow, seEndCol)).Top
        .Width = Sheets(bushomei).Range(Cells(seStartRow, seStartCol), Cells(seEndRow, seEndCol)).Width
        .Height = Sheets(bushomei).Range(Cells(seStartRow, seStartCol), Cells(seEndRow, seEndCol)).Height
        .Placement = xlMoveAndSize
    End With
    
    If userNumber Mod 4 = 0 Then
        Sheets(bushomei).HPageBreaks.Add Rows(kyuEndRow + 2)
    End If

End Sub

