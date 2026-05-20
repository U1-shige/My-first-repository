' ============================================================
' 複数Excelファイルを1つに集約するマクロ
'
' 使い方:
'   1. このファイルの内容を集約用.xlsm の標準モジュールに貼り付ける
'   2. AggregateExcel を実行する
'   3. ダイアログで作業Excelが入ったフォルダを選択する
'   4. 同じブックの「集約データ」シートに結果が出力される
' ============================================================

Sub AggregateExcel()

    Dim targetFolder As String
    Dim outputSheet  As Worksheet
    Dim headerWritten As Boolean
    Dim nextRow      As Long

    ' ── フォルダ選択ダイアログ ──────────────────────────
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "作業Excelが入ったフォルダを選択してください"
        If .Show <> True Then
            MsgBox "キャンセルされました。", vbInformation
            Exit Sub
        End If
        targetFolder = .SelectedItems(1) & "\"
    End With

    ' ── 出力シートの準備 ────────────────────────────────
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("集約データ")
    On Error GoTo 0

    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        outputSheet.Name = "集約データ"
    Else
        outputSheet.Cells.Clear
    End If

    headerWritten = False
    nextRow = 1

    ' ── フォルダ内の .xlsx / .xls を処理 ──────────────
    Dim fileName As String
    Dim srcBook  As Workbook
    Dim srcSheet As Worksheet
    Dim lastRow  As Long
    Dim lastCol  As Long
    Dim fileCount As Long
    Dim rowCount  As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fileName = Dir(targetFolder & "*.xlsx")
    Do While fileName <> ""
        ' 自分自身はスキップ
        If fileName <> ThisWorkbook.Name Then
            Set srcBook = Workbooks.Open(targetFolder & fileName, ReadOnly:=True)
            Set srcSheet = srcBook.Sheets(1)  ' 先頭シートを対象

            lastRow = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
            lastCol = srcSheet.Cells(1, srcSheet.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 And lastCol >= 1 Then
                ' 最初のファイルだけヘッダーをコピー（「出所ファイル」列を先頭に追加）
                If Not headerWritten Then
                    outputSheet.Cells(nextRow, 1).Value = "出所ファイル"
                    srcSheet.Range(srcSheet.Cells(1, 1), srcSheet.Cells(1, lastCol)) _
                        .Copy outputSheet.Cells(nextRow, 2)
                    nextRow = nextRow + 1
                    headerWritten = True
                End If

                ' データ行をコピー（ヘッダー行=1行目を除く）
                If lastRow >= 2 Then
                    Dim dataRows As Long
                    dataRows = lastRow - 1
                    outputSheet.Cells(nextRow, 1).Resize(dataRows, 1).Value = fileName
                    srcSheet.Range(srcSheet.Cells(2, 1), srcSheet.Cells(lastRow, lastCol)) _
                        .Copy outputSheet.Cells(nextRow, 2)
                    nextRow = nextRow + dataRows
                    rowCount = rowCount + dataRows
                End If
            End If

            srcBook.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If

        fileName = Dir()
    Loop

    ' ── 列幅の自動調整 ──────────────────────────────────
    outputSheet.Columns.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If fileCount = 0 Then
        MsgBox "対象のExcelファイルが見つかりませんでした。" & vbCrLf & targetFolder, vbExclamation
    Else
        MsgBox fileCount & " ファイル / " & rowCount & " 行を集約しました。" & vbCrLf & _
               "シート「集約データ」を確認してください。", vbInformation
    End If

End Sub
