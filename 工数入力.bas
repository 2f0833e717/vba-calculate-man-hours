Option Explicit

Private Sub TransferData()
  Dim wsSource As Worksheet, wsTarget As Worksheet
  Dim lastRow As Long, i As Long, targetRow As Long
  Dim yyyymm As String, yyyymmdd As String, workName As String, subWorkName2 As String
  Dim time As Double
  Dim timeDict As Object, key As String
  Dim colOffset As Interger

  'ソースワークしーーとを設定
  Set wsSource = ActiveWorkbook.Sheets("Sheet1")

  'ソースワークシートの最終行を取得
  lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Rows

  yyyymm = Mid(wsSource.Cells(2, "A").Value, 6, 6)

  Dim ws As Worksheet
  On Error Resume Next
  Set ws = Sheets(yyyymm)
  On Error GoTo 0

  If ws.Name = ActiveSheet.Name Then
    Sheets(yyyymm).Select
    If Not IsEmpty(Sheets(yyyymm).Range("G9").Value) Then
      Sheets(yyyymm).Cells.ClearCounts
    End If
  Else
    Exit Sub
  End If

  'ヘッダーが１行目にあると仮定
  For i = 2 To lastRow
    'A列のfileNameからyyyymmとyyyymmddを抽出
    yyyy = Mid(wsSource.Cells(i, "A").Value, 6, 6)
    yyyymm = Mid(wsSource.Cells(i, "A").Value, 6, 8)
    
    'ターゲットワークシートが存在するかチェック、存在しない場合は作成
    On Error Resume Next
    Set wsTarget = ActiveWorkbook.SHeets(yyyymm)
    If wsTarget Is Nothing Then
      Set wsTarget = Sheets.Add(After:=Sheets(Sheets.Count))
      wsTarget.Name = yyyymm
    End If
    On Error GoTo 0

    'ソースワークシートから値を取得
    workName = wsSource.Cells(i, "E").Value
    subWorkName2 = wsSource.Cells(i, "G").Value
    time = wsSource.Cells(i, "D").Value

    '列オフセットを計算（日付が変わるたびに列が変わる）
    colOffset = Day(CDate(Mid(yyyymmdd, 5, 2) & "/" _
      & Right(yyyymmdd, 2) & "/" & Left(yyyymmdd, 4)))
    
    'yyyymmddの値を日付型のyyyy/mm/dd形式に変換
    Dim dt As Date
    dt = DateSerial(Left(yyyymmdd, 4), Mid(yyyymmdd, 5, 2), Right(yyyymmdd, 2))

    'yyyy/mm/dd形式の日付をヘッダーとして設定
    wsTarget.cell(8, colOffset + 8).Value = Format(dt, "yyyy/mm/dd")

    'sum関数をヘッダーとして設定
    '最終行（MaxRow）を取得
    MaxRow = wsTarget.cell(wsTarget.Rows.Count, "G").End(xlUp).Row

    '最終列を取得
    Dim lastCol As Long
    lastCol = wsTarget.cell(wsTarget.Columns.Count).End(xlToLeft).Columns

    '各列の合計を計算し、ヘッダー行に表示
    Dim j As Long
    For j = 9 To lastCol  
      wsTarget.cell(7, j).formula = "=SUM(" & wsTarget.Cells(9, j).Address & ":" & wsTarget.Cells(MaxRow, j).Address & ")"
    Next j

    'ターゲットワークシートでデータ入力のための次の空行を見つける
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "G").End(xlUp).Row + 1
    'データ入力は９行目から開始
    If targetRow < 9 Then targetRow = 9

    'ターゲットワークシートでworkNameがすでに存在するかチェック
    Dim rng As Range
    Set rng = ws Target.Range("G9:G" & targetRow).Find(workName)

    If rng Is Nothing Then
      'workNameが存在しない場合、新しいデータを入力
      wsTarget.Cells(targetRow, "G").Value = workName
      wsTarget.Cells(targetRow, 7).Value = wsTarget.Cells(targetRow, 7).Value & subWorkName2
      wsTarget.Cells(targetRow, colOffset + 8).Value = time
      '合計timeをDictionaryに保存
      key = yyyymmdd & "|" & workName
      timeDict(key) = time
    Else
      'workNameが存在する場合、既存のデータを更新
      '同じworkNameに紐づいたsubWorkName2がすでに存在する場合は、改行して追記
      If InStr(wsTarget.Cells(rng.Row, colOffset + 7).Value, subWorkName2) = 0 Then
        wsTarget.Cells(rng.Row, 7).Value = wsTarget.Cells(rng.Row, 7).Value & vbCrLf & subWorkName2
      End If
      '同じ日付のworkNameのtimeの場合は合算されて当日分の合計値を同じ列に入力
      key = yyyymmdd & "|" & workName
      If timeDict.Exists(key) Then
        timeDict(key) = timeDict(key) + time
      Else
        timeDict(key) = time
      End If
      wsTarget.Cells(rng.Row, colOffset + 8).Value = timeDict(key)
    End If
  Next i

  '最大行数取得
  Dim MaxRow As Long
  MaxRow = ActiveWorkbook.ActiveSheet.Cells(10000, 7).End(xlUp).Row
  Row("8:" & MaxRow).Select
  Selection.RowHeight = 18.75
  Range("A1").Select
End Sub