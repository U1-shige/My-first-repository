' ============================================================
' 作業ファイルのR・S・T列を、C列の値をキーに集計ファイルへ転記するマクロ
'
' 使い方:
'   1. このコードを集計ファイル(.xlsm) の標準モジュールに貼り付ける
'   2. PasteRSTColumns を実行する
'   3. ダイアログ①: 作業ファイルが入ったフォルダを選択
'   4. ダイアログ②: 集計ファイル(.xlsx/.xlsm) を選択
'   5. 完了メッセージで転記件数を確認する
'
' 仕様:
'   - 作業ファイルのC列と集計ファイルのC列を比較してマッチした行に転記
'   - 作業ファイルのR・S・T列の値を集計ファイルの対応行へ貼り付け
'   - C列に重複値(例:597)がある場合は両行に転記し、ログシートに記録
' ============================================================

Option Explicit

Sub PasteRSTColumns()

    ' ── ① フォルダ選択 ────────────────────────────────
    Dim workFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "作業ファイルが入ったフォルダを選択"
        If .Show <> True Then MsgBox "キャンセルされました。", vbInformation: Exit Sub
        workFolder = .SelectedItems(1) & "\"
    End With

    ' ── ② 集計ファイル選択 ────────────────────────────
    Dim summaryPath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "集計ファイルを選択"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
        If .Show <> True Then MsgBox "キャンセルされました。", vbInformation: Exit Sub
        summaryPath = .SelectedItems(1)
    End With

    ' ── ③ 集計ファイルを開く ──────────────────────────
    Dim sumBook  As Workbook
    Dim sumSheet As Worksheet
    Set sumBook  = Workbooks.Open(summaryPath)
    Set sumSheet = sumBook.Sheets(1)

    ' 集計ファイルのC列をDictionaryに読み込む（値 → 行番号リスト）
    Dim sumMap As Object
    Set sumMap = CreateObject("Scripting.Dictionary")

    Dim sumLastRow As Long
    sumLastRow = sumSheet.Cells(sumSheet.Rows.Count, 3).End(xlUp).Row

    Dim r As Long
    For r = 2 To sumLastRow   ' 1行目はヘッダーとして除外
        Dim keyVal As String
        keyVal = Trim(CStr(sumSheet.Cells(r, 3).Value))
        If keyVal <> "" Then
            If sumMap.Exists(keyVal) Then
                ' 重複キー: 既存リストに行番号を追加
                Dim arr() As Long
                arr = sumMap(keyVal)
                ReDim Preserve arr(UBound(arr) + 1)
                arr(UBound(arr)) = r
                sumMap(keyVal) = arr
            Else
                Dim newArr(0) As Long
                newArr(0) = r
                sumMap(keyVal) = newArr
            End If
        End If
    Next r

    ' ── ④ ログシートの準備 ────────────────────────────
    Dim logSheet As Worksheet
    On Error Resume Next
    Set logSheet = sumBook.Sheets("転記ログ")
    On Error GoTo 0
    If logSheet Is Nothing Then
        Set logSheet = sumBook.Sheets.Add(After:=sumBook.Sheets(sumBook.Sheets.Count))
        logSheet.Name = "転記ログ"
    Else
        logSheet.Cells.Clear
    End If
    logSheet.Range("A1:E1").Value = Array("作業ファイル", "C列キー", "集計行", "ステータス", "備考")
    Dim logRow As Long
    logRow = 2

    ' ── ⑤ 作業ファイルを順番に処理 ───────────────────
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim fileName  As String
    Dim fileCount As Long
    Dim matchCount As Long
    Dim noMatchCount As Long

    fileName = Dir(workFolder & "*.xlsx")
    Do While fileName <> ""

        Dim wkBook  As Workbook
        Dim wkSheet As Worksheet
        Set wkBook  = Workbooks.Open(workFolder & fileName, ReadOnly:=True)
        Set wkSheet = wkBook.Sheets(1)

        Dim wkLastRow As Long
        wkLastRow = wkSheet.Cells(wkSheet.Rows.Count, 3).End(xlUp).Row

        Dim i As Long
        For i = 2 To wkLastRow   ' 1行目ヘッダー除外
            Dim wkKey As String
            wkKey = Trim(CStr(wkSheet.Cells(i, 3).Value))
            If wkKey = "" Then GoTo NextRow

            ' R・S・T列の値を取得
            Dim vR As Variant, vS As Variant, vT As Variant
            vR = wkSheet.Cells(i, 18).Value  ' R列
            vS = wkSheet.Cells(i, 19).Value  ' S列
            vT = wkSheet.Cells(i, 20).Value  ' T列

            If sumMap.Exists(wkKey) Then
                Dim rows() As Long
                rows = sumMap(wkKey)
                Dim isDuplicate As Boolean
                isDuplicate = (UBound(rows) > 0)

                Dim j As Long
                For j = 0 To UBound(rows)
                    Dim destRow As Long
                    destRow = rows(j)

                    ' 集計ファイルのR・S・T列を上書き
                    ' ※集計ファイルの列位置は作業ファイルと同じR=18,S=19,T=20と仮定
                    '   異なる場合は下記の列番号を変更してください
                    sumSheet.Cells(destRow, 18).Value = vR
                    sumSheet.Cells(destRow, 19).Value = vS
                    sumSheet.Cells(destRow, 20).Value = vT

                    Dim note As String
                    note = IIf(isDuplicate, "重複キー(" & wkKey & ") - " & (j + 1) & "行目に転記", "")
                    logSheet.Cells(logRow, 1).Value = fileName
                    logSheet.Cells(logRow, 2).Value = wkKey
                    logSheet.Cells(logRow, 3).Value = destRow
                    logSheet.Cells(logRow, 4).Value = "転記済"
                    logSheet.Cells(logRow, 5).Value = note
                    logRow = logRow + 1
                    matchCount = matchCount + 1
                Next j
            Else
                ' キーが見つからなかった行をログに記録
                logSheet.Cells(logRow, 1).Value = fileName
                logSheet.Cells(logRow, 2).Value = wkKey
                logSheet.Cells(logRow, 3).Value = "-"
                logSheet.Cells(logRow, 4).Value = "未一致"
                logSheet.Cells(logRow, 5).Value = "集計ファイルにキーなし"
                logRow = logRow + 1
                noMatchCount = noMatchCount + 1
            End If

NextRow:
        Next i

        wkBook.Close SaveChanges:=False
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    ' ── ⑥ 集計ファイルを保存して閉じる ──────────────
    logSheet.Columns.AutoFit
    sumBook.Save
    sumBook.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "完了しました。" & vbCrLf & vbCrLf & _
           "処理ファイル数 : " & fileCount & " 件" & vbCrLf & _
           "転記成功       : " & matchCount & " 行" & vbCrLf & _
           "未一致(スキップ): " & noMatchCount & " 行" & vbCrLf & vbCrLf & _
           "詳細は「転記ログ」シートを確認してください。", vbInformation

End Sub
