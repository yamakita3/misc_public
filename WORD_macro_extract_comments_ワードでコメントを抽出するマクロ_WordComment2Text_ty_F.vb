' 全体での行番号を取得する関数を追加
' https://akashi-keirin.hatenablog.com/entry/2020/05/31/212239
Public Function getLineNumber(ByVal tgtRange As Range) As Long
  Dim ret As Long
  Dim currPage As Long
  currPage = tgtRange.Information(wdActiveEndPageNumber)  'tgtRangeのあるページ番号を取得'
  Dim currLine As Long
  currLine = tgtRange.Information(wdFirstCharacterLineNumber)  'tgtRangeのあるページ内での行番号を取得'
  If currPage = 1 Then  'tgtRangeが1ページ目にあるときは、その行番号を返す'
    ret = currLine
    GoTo Finalizer:
  End If   '2ページ以上ある時は、手前のページまでの累計を足さなければいけない'
  Dim Doc As Document  
  Set Doc = tgtRange.Parent'親ドキュメントオブジェクトを取得'
  Dim orgRange As Range
  Set orgRange = Selection.Range  'カーソル位置を記録'
  Call Doc.Range(0, 0).Select  '文書の先頭にカーソルを置く'
  Dim pageEnd As Long
  '1ページ目の最終位置を選択'
  Dim i As Long
  For i = 1 To currPage - 1
    pageEnd = Doc.Bookmarks("\Page").End    'ページの最終位置を取得'
    Call Doc.Range(pageEnd - 1, pageEnd - 1).Select    'ページの末尾にカーソルを置く'
    ret = ret + Selection.Range.Information(wdFirstCharacterLineNumber)    'ページ末尾の行番号＝そのページの総行数を加算'
    Call Selection.MoveRight(wdCharacter, 1, wdMove)    '次のページの先頭へ'
  Next
  ret = ret + currLine
  Call orgRange.Select  'カーソル位置を戻す'
  Finalizer:
  getLineNumber = ret
End Function

Sub WordComment2Text_コメントをテキストに書き出すマクロ()
'参考資料 https://www.wordvbalab.com/code/5853/
    Dim doc As Document
    Dim comment As Comment
    Dim outputFile As String
    Dim fso As Object
    Dim ts As Object
    Dim pageNum As Long
    Dim lineNum As Long
    Dim totalLineNum As Long ' 追加: 全体での行番号
    Dim sentenceText As String
    Dim i As Long
    Dim maxSentenceLength As Long ' 追加: センテンスの最大長さ
    
    Set doc = ActiveDocument' アクティブなドキュメントを設定
    
    ' コメントがない場合は終了
    If doc.Comments.Count = 0 Then
        MsgBox "ドキュメントにコメントがありません。", vbInformation
        Exit Sub
    End If
    
    ' ユーザーに確認
    'If MsgBox("ファイル内のすべてのコメントを書き出しますか？", _
    '    vbQuestion Or vbYesNo, "実施前の確認") = vbNo Then
    '    Exit Sub
    'End If

   'ユーザーにセンテンスの最大長さを入力してもらう
    maxSentenceLength = InputBox("リード文の最大文字数（0で処理中止）", "設定", "20")
    If maxSentenceLength <= 0 Then
        MsgBox "無効な入力です。処理を中止します。", vbExclamation
        Exit Sub
    End If  
  
    ' 出力ファイルの選択
    Set dlgSave = Application.FileDialog(msoFileDialogSaveAs)
    With dlgSave
    ' With Application.FileDialog(msoFileDialogSaveAs)
    '' Application.FileDialog(msoFileDialogSaveAs)を使った場合、ファイルフィルタが使えない。
        .Title = "コメント出力ファイルの保存"
        .FilterIndex = 13 ' テキストファイルを選択
        '.Filters.Clear
        '.Filters.Add "テキストファイル", "*.txt"
        '.Filters.Add "すべてのファイル", "*.*"
        .InitialFileName = "comments_output.txt"
        If .Show = -1 Then
            outputFile = .SelectedItems(1)
        Else
            MsgBox "ファイルの保存がキャンセルされました。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' ファイル書き込みの準備
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(outputFile, True)
    
    ' ヘッダー（列名）を書き込む
    ts.WriteLine "全体行番号,ページ番号,ページ内行番号,作成者,本文のセンテンス,選択テキスト,コメントテキスト"
    
    ' 各コメントを処理
    For i = 1 To doc.Comments.Count
        With doc.Comments(i)
            pageNum = .Scope.Information(wdActiveEndPageNumber) '(wdActiveEndAdjustedPageNumber)との違いは？
            lineNum = .Scope.Information(wdFirstCharacterLineNumber)
            totalLineNum = getLineNumber(.Scope) ' 全体での行番号を取得
            ' センテンスを取得し、必要に応じて省略
            sentenceText = .Scope.Sentences(1).Text
            If Len(sentenceText) > maxSentenceLength Then
                sentenceText = Left(sentenceText, maxSentenceLength) & "..."
            End If
            sentenceText = Replace(sentenceText, ",", "，") ' カンマをエスケープ
            sentenceText = Replace(sentenceText, vbCr, "")' 改行を削除
            sentenceText = Replace(sentenceText, vbLf, "")
            
            ' ファイルに書き込む
            ts.WriteLine  totalLineNum & "," & pageNum & "," & lineNum & "," & .Author & ","  & sentenceText & ","  & Replace(.Scope.Text, ",", "，") & "," & Replace(.Range.Text, vbCr, "")'Replaceは何故必要？
        End With
    Next i
        
    ts.Close' ファイルを閉じる
    
    ' 完了メッセージ
    MsgBox "コメントの抽出が完了しました。" & vbNewLine & "保存先: " & outputFile, vbInformation

    ' オブジェクトの解放
    Set ts = Nothing
    Set fso = Nothing
    Set doc = Nothing
End Sub