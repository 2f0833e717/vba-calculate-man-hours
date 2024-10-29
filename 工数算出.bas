Option Explicit

Sub ReadTextFiles()
  Dim folderPath As String
  
  '指定のフォルダパスを設定
  folderPath = "C:\Users\user\work\000_memo\_bk\yyyymm"

  Application.DisplayAlerts = False

  Dim ws As WorkSheet
  On Error Resume Next
  Set ws = Sheets("Sheet1")
  On Error GoTo 0
  If ws.Name = ActiveSheet.Name Then
    Sheets("Sheet1").Select
    If Not IsEmpty(Sheets("Sheet1").Range("A1").Value) Then
      Sheets("SHeet1").Cells.ClearContents
    End If
  Else
    Exit Sub
  End If

  '指定されたフォルダ内のテキストファイルを読み込む
  ReadTextFilesInFolder folderPath
  
  Application.DisplayAlerts = True

  Call AddHeader
  Call SheetPaste

  MsgBox ("success")
End Sub

Private Sub ReadTextFilesInFolder(folderPath As String)
  Dim fileName As String
  Dim fileContent As String
  Dim startTimestamp As String
  Dim endTimestamp As String
  Dim subWorkName As String
  Dim workTime As String
  Dim i As Long

  'フォルダ内の最初のファイルを取得
  fileName = Dir(folderPath & "\*", vbDirectory)

  'ファイルが無くなるまでループ
  Do While fileName <> ""
    If (GetAttr(folderPath & "\" & fileName) And vbDirectory) = vbDirectory Then
      If fileName <> "." And fileName <> ".." Then
        'ディレクトリの場合、再帰的に探索
        ReadTextFilesInFolder folderPath & "\" & fileName
      End If
    ElseIf Right(fileName, 4) = ".txt" Then
      'テキストファイルを開き、内容を読み込む
      fileContent = GetFileContent(folderPath & "\" & fileName)

      'ファイル内容を解析
      Do While Len(fileContent) > 0
        startTimestamp = ""
        mainWorkName = ""
        subWorkName = ""
        endTimestamp = ""
        
        'startTimestamp
        fileContent = GetTimestamp(fileContent, startTimestamp)

        If startTimestamp <> "" Then
          
          'mainWorkName
          fileContent = GetWorkName(fileContent, mainWorkName)

          'subWorkName
          If mainWorkName = "内職" Then
            fileContent = GetWorkName(fileContent, subWorkName)
          End If

          'endTimestamp
          fileContent = GetTimestamp(fileContent, endTimestamp)
          If endTimestamp <> "" Then
            workTime = GetWorkTime(startTimestamp, endTimestamp)
          Else
            workTime ""
          End If

          'Excelの表に表示
          If IsNumeric(workTime) Then
            If CLng(workTime) < 1000 Then
              i = i + 1
              Cells(1, 1).Value = fileName
              Cells(1, 2).Value = startTimestamp
              Cells(1, 3).Value = endTimestamp
              Cells(1, 4).Value = workTime
              If Left(Trim(mainWorkName), 1) = "=" Then
                Cells(1, 5).Value = "'" & Trim(mainWorkName)
              Else
                Cells(1, 5).Value = Trim(mainWorkName)
              End If
              If Left(Trim(subWorkName), 1) = "=" Then
                Cells(1, 6).Value = "'" & Trim(subWorkName)
              Else
                Cells(1, 6).Value = Trim(subWorkName)
              End If
            End If
          End If
        Else
          Exit Do
        End If
      Loop
    End If

    '次のファイルを取得
    fileName = Dir
  Loop
End Sub

Function GetFileContent(filePath As String) As String
  Dim fileNumber As Interger
  Dim fileContent As String
  Dim textLine As String

  fileNumber = FreeFile
  Open filePath For Input As fileNumber
  Do Until EOF(fileNumber)
    Line Input #fileNumber, textLine
    fileContent = fileContent & textLine & vbCrLf
  Loop
  Close fileNumber

  getFileContent = fileContent
End Function

Function GetTimestamp(fileContent As String, ByRef timestamp As String) As String
  Dim regex As Object
  Dim matches As Object

  Set regex = CreateObject("VBScript.RegExp")
  regex.Pattern  "(\d{4}年\d{1,2}月\d{1,2}日　\d{1,2}:\d{2})"
  regex.Global = False
  'UTF-8保存だと文字化けしてマッチングしないので注意
  Set matches = regex.Execute(fileContent)

  If matches.Count > 0 Then
    timestamp = matches(0) Then
    'fileContent = Mid(fileContent, matches(0).FirstIndex + Len(timestamp) + 2)
  Else
    timestamp = ""
  End If

  GetTimestamp = fileContent
End Function

Function GetWorkName(fileContent As String, ByRef workName As String) As String
  Dim regex As Object
  Dim pos As Interger
  Dim matches As Object

  Set regex = CreateObject("VBScript.RegExp")
  regex.Pattern  "(\d{4}年\d{1,2}月\d{1,2}日　\d{1,2}:\d{2})"
  regex.Global = False
  Set matches = regex.Execute(fileContent)

  workName = ""
  If matches.Count > 0 Then
    Do While (workName = "" _
      Or regex.Test(workName) _
      Or workName = vbCrLf _
      Or Trim(workName) = "" _
      Or Trim(workName) = "　") _
      And Len(fileContent) > 0
      '改行の有無を確認
      pos = InStr(fileContent, vbCrLf)
      If pos > 0 Then
        '改行以外の文字列を取得し格納
        workName = Left(fileContent, pos -1)
        '該当の行を削除、削除後のテキストデータを変数に格納
        fileContent = Mid(fileContent, pos + Len(vbCrLf))
      Else
        Exit Do
      End If
    Loop
  Else

  End If

  GetWorkName = fileContent
End Function

Function GetWorkNameSub(fileContent As String, ByRef workName As String) As String
  Dim regex As Object
  Dim pos As Interger
  Dim matches As Object

  Set regex = CreateObject("VBScript.RegExp")
  regex.Pattern  "(\d{4}年\d{1,2}月\d{1,2}日　\d{1,2}:\d{2})"
  regex.Global = False
  Set matches = regex.Execute(fileContent)

  workName = ""
  If matches.Count > 0 Then
    Do While (workName = "" _
      Or regex.Test(workName) _
      Or workName = vbCrLf _
      Or Trim(workName) = "" _
      Or Trim(workName) = "　") _
      And Len(fileContent) > 0
      '改行の有無を確認
      pos = InStr(fileContent, vbCrLf)
      If pos > 0 Then
        '改行以外の文字列を取得し格納
        workName = Left(fileContent, pos -1)
        '該当の行を削除、削除後のテキストデータを変数に格納
        fileContent = Mid(fileContent, pos + Len(vbCrLf))
      Else
        Exit Do
      End If
    Loop
  Else

  End If

  GetWorkName = fileContent
End Function

Function GetWorkTime(startTimestamp As String, endTimestamp As String) As String
  Dim startTime As Date
  Dim endTime As Date
  Dim workTime As String

  If Len(startTimestamp) > 0 Then
    startTime = CDate(Replace(Replace(Replace(Replace(startTimestamp, "年", "/"), "月", "/"), "日", ""), "　", " "))
    If Len(endTimestamp) > 0 Then
      endTime = CDate(Replace(Replace(Replace(Replace(startTimestamp, "年", "/"), "月", "/"), "日", ""), "　", " "))
      workTime = DateDiff("n", startTime,endTime)
    Else
      workTime  ""
    End If
  Else
    workTime = ""
  End If

  GetWorkTime = workTime
End Function

Private Sub AddHeader()

  Sheet("Sheet1").Select

  Rows("1:1").Select
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

  Range("A1").Select
  ActiveCell.FormulaR1C1 = "fileName"
  Range("B1").Select
  ActiveCell.FormulaR1C1 = "start"
  Range("C1").Select
  ActiveCell.FormulaR1C1 = "end"
  Range("D1").Select
  ActiveCell.FormulaR1C1 = "time"
  Range("E1").Select
  ActiveCell.FormulaR1C1 = "workName"
  Range("F1").Select
  ActiveCell.FormulaR1C1 = "subWorkName"
  Range("G1").Select
  ActiveCell.FormulaR1C1 = "subWorkName2"

  If Not ActiveSheet.AutoFilterMode Then
    Range("A1:G1").AutoFilter
  End If

  Dim rng As Range
  Dim cell As Range
  Dim filterArray() As Variant
  Dim i As Interger
  Dim lastRorw As Long

  lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Rows
  Set rng = ActiveSheet.Range("D2:D" & lastRorw)

  i = 0
  For Each cell In rng
    If Cell.Value <> 0 Then
      ReDim Preserve filterArray(i)
      filterArray(i) = cell.Value
      i = i + 1
    End If
  Next cell

  If i > 0 Then
    rng.AutoFilter Field:=4, Criteria1:="<>" & 0
  End If
End Sub

Private Sub SheetPaste()
  Dim ws As WorkSheet
  Dim rng As Range

  'フィルタリングされたデータをコピー
  Set rng = ActiveWorkbook.Sheets("Sheet1").AutoFilter.Range.SpecialCells(xlCellTypeVisible)
  rng.Copy

  '新しいシートを作成し、データを貼り付け
  Set ws = ActiveWorkbook.Sheets.Add
  ws.Paste

  '元のシートを削除
  Application.DisplayAlerts = False
  ActiveWorkbook.Sheets("Sheet1").Delete
  Application.DisplayAlerts = True

  '新しいシートの名前を「Sheet1」に変更
  ws.Name = "Sheet1"
End Sub