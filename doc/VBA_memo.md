# Word VBA マクロのメモ

## まずはリファレンス
### [WdColorIndex列挙(Word)]
> 適用される色を指定します。

<details>
<summary>📝 WdColorIndex 列挙値の一覧：</summary>

| 名前          | 値 | 説明                                 |
|---------------|----|--------------------------------------|
| wdAuto        | 0  | 自動設定。 通常の既定値は黒です。    |
| wdBlack       | 1  | 黒                                   |
| wdBlue        | 2  | 青                                   |
| wdBrightGreen | 4  | 明るい緑                             |
| wdByAuthor    | -1 | 文書の作成者が定義した色             |
| wdDarkBlue    | 9  | 濃い青                               |
| wdDarkRed     | 13 | 濃い赤                               |
| wdDarkYellow  | 14 | 濃い黄                               |
| wdGray25      | 16 | 網かけ 25 の灰色                     |
| wdGray50      | 15 | 網かけ 50 の灰色                     |
| wdGreen       | 11 | 緑                                   |
| wdNoHighlight | 0  | 適用されている強調表示を解除します。 |
| wdPink        | 5  | ピンク                               |
| wdRed         | 6  | 赤                                   |
| wdTeal        | 10 | 青緑                                 |
| wdTurquoise   | 3  | 水色                                 |
| wdViolet      | 12 | 紫                                   |
| wdWhite       | 8  | 白                                   |
| wdYellow      | 7  | 黄                                   |

</details>

### [WdFindWrap列挙(Word)]
> 検索対象の選択範囲または指定範囲内に検索文字列が見つからなかった場合の、折り返し動作を指定します。

<details>
<summary>📝 WdFindWrap 列挙値の一覧：</summary>

| 名前           | 値 | 説明                                                                                                 |
|----------------|----|------------------------------------------------------------------------------------------------------|
| wdFindAsk      | 2  | 選択範囲または指定範囲を検索し、文書の残りの部分も検索するかどうかをたずねるメッセージを表示します。 |
| wdFindContinue | 1  | 検索範囲の先頭または末尾まで検索し、さらに検索を続けます。                                           |
| wdFindStop     | 0  | 検索範囲の先頭または末尾まで検索したら、検索を終了します。                                           |

</details>

### [WdBuiltInProperty列挙(Word)]
> あらかじめ用意されているドキュメントプロパティを指定します。

<details>
<summary>📝 WdBuiltInProperty 列挙値の一覧：</summary>

| 名前                      | 値 | 説明                     |
|---------------------------|----|--------------------------|
| wdPropertyAppName         | 9  | アプリケーションの名前   |
| wdPropertyAuthor          | 3  | 作成者                   |
| wdPropertyBytes           | 22 | バイト数                 |
| wdPropertyCategory        | 18 | 分類                     |
| wdPropertyCharacters      | 16 | 文字数                   |
| wdPropertyCharsWSpaces    | 30 | 文字数 (スペースを含む)  |
| wdPropertyComments        | 5  | コメント                 |
| wdPropertyCompany         | 21 | 会社名                   |
| wdPropertyFormat          | 19 | サポートされていません。 |
| wdPropertyHiddenSlides    | 27 | サポートされていません。 |
| wdPropertyHyperlinkBase   | 29 | サポートされていません。 |
| wdPropertyKeywords        | 4  | キーワード               |
| wdPropertyLastAuthor      | 7  | 最終作成者               |
| wdPropertyLines           | 23 | 行数                     |
| wdPropertyManager         | 20 | 管理者                   |
| wdPropertyMMClips         | 28 | サポートされていません。 |
| wdPropertyNotes           | 26 | メモ                     |
| wdPropertyPages           | 14 | ページ数                 |
| wdPropertyParas           | 24 | 段落数                   |
| wdPropertyRevision        | 8  | 改訂番号                 |
| wdPropertySecurity        | 17 | セキュリティ設定         |
| wdPropertySlides          | 25 | サポートされていません。 |
| wdPropertySubject         | 2  | 副題                     |
| wdPropertyTemplate        | 6  | テンプレート名           |
| wdPropertyTimeCreated     | 11 | 作成日時                 |
| wdPropertyTimeLastPrinted | 10 | 最終印刷日時             |
| wdPropertyTimeLastSaved   | 12 | 最終更新日時             |
| wdPropertyTitle           | 1  | タイトル                 |
| wdPropertyVBATotalEdit    | 13 | VBA プロジェクトの編集数 |
| wdPropertyWords           | 15 | 単語数                   |

</details>

### [WdRemoveDocInfoType列挙(Word)]
> 文書から削除する情報の種類を指定します。

<details>
<summary>📝 WdRemoveDocInfoType 列挙値の一覧：</summary>

| 名前                           | 値 | 説明                                                     |
|--------------------------------|----|----------------------------------------------------------|
| wdRDIAll                       | 99 | すべての文書情報を削除します。                           |
| wdRDIComments                  | 1  | 文書のコメントを削除します。                             |
| wdRDIContentType               | 16 | コンテンツ タイプの情報を削除します。                    |
| wdRDIDocumentManagementPolicy  | 15 | ドキュメント管理ポリシーの情報を削除します。             |
| wdRDIDocumentProperties        | 8  | 文書プロパティを削除します。                             |
| wdRDIDocumentServerProperties  | 14 | ドキュメント サーバーのプロパティを削除します。          |
| wdRDIDocumentWorkspace         | 10 | ドキュメント ワークスペースの情報を削除します。          |
| wdRDIEmailHeader               | 5  | 電子メール ヘッダー情報を削除します。                    |
| wdRDIInkAnnotations            | 11 | インク注釈を削除します。                                 |
| wdRDIRemovePersonalInformation | 4  | 個人情報を削除します。                                   |
| wdRDIRevisions                 | 2  | 変更履歴マークを削除します。                             |
| wdRDIRoutingSlip               | 6  | 回覧先情報を削除します。                                 |
| wdRDISendForReview             | 7  | 校閲者に文書を送信するときに格納された情報を削除します。 |
| wdRDITemplate                  | 9  | テンプレート情報を削除します。                           |
| wdRDITaskpaneWebExtensions     | 17 | 作業ウィンドウの web 拡張機能の情報を削除します。        |
| wdRDIVersions                  | 3  | 文書のバージョン情報を削除します。                       |

</details>

* マクロの例
    <details>
    <summary>🎨 ドキュメントの文書プロパティを削除する:</summary>

    ドキュメントの文書プロパティを削除します。**※ファイルのプロパティに限る模様…**
    ```vb
    Sub ドキュメントの文書プロパティを削除する()
        Dim objDoc As Document
        Set objDoc = ActiveDocument
        
        Call objDoc.RemoveDocumentInformation(wdRDIDocumentProperties)
        
        Set objDoc = Nothing
    End Sub
    ```
    
    </details>

### [Document.RemovePersonalInformationプロパティ(Word)]
> True 場合は、Microsoft Word は、コメント、変更履歴、およびドキュメントの保存時に [プロパティ] ダイアログ ボックスからすべてのユーザー情報を削除します。 読み取り/書き込みが可能な Boolean です。

* マクロの例
    <details>
    <summary>🎨 文書を上書き保存するときに個人情報が削除されるように設定:</summary>

    ユーザーが次に現在の文書を保存したときに、この文書から個人情報が削除されるように設定します。
    ```vb
    Sub 文書を上書き保存するときに個人情報が削除されるように設定()
        Dim objDoc As Document
        Set objDoc = ActiveDocument
        
        objDoc.RemovePersonalInformation = True

        Set objDoc = Nothing
    End Sub
    ```

    </details>

### [WdInformation列挙(Word)]
> 指定された選択範囲または指定範囲に関して取得される情報の種類を指定します。

<details>
<summary>📝 WdInformation 列挙値の一覧：</summary>

| 名前                                       | 値 | 説明                                                                                                                                                                                                                                                              |
|--------------------------------------------|----|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| wdActiveEndAdjustedPageNumber              | 1  | 指定された選択範囲または指定範囲のアクティブな終点が含まれるページの数を返します。 開始ページ番号を設定した場合、または他の手動調整を行う場合は、調整済みページ番号 ( wdActiveEndPageNumber とは異なる) を返します。                                              |
| wdActiveEndPageNumber                      | 3  | 指定された選択範囲または文書の先頭から数えて、範囲のアクティブな終点が含まれるページの数を返します。 ページ番号の手動調整は無視されます ( wdActiveEndAdjustedPageNumber とは異なり)。                                                                             |
| wdActiveEndSectionNumber                   | 2  | 指定された選択範囲または指定範囲の終了位置を含むセクション番号を取得します。                                                                                                                                                                                      |
| wdAtEndOfRowMarker                         | 31 | 指定された選択範囲または指定範囲が表の中の行区切り記号である場合、値は True です。                                                                                                                                                                                |
| wdCapsLock                                 | 21 | Returns True if Caps Lock is in effect.                                                                                                                                                                                                                           |
| wdEndOfRangeColumnNumber                   | 17 | 指定された選択範囲または指定範囲の終了位置の列番号を取得します。                                                                                                                                                                                                  |
| wdEndOfRangeRowNumber                      | 14 | 指定された選択範囲または指定範囲の終了位置の行番号を取得します。                                                                                                                                                                                                  |
| wdFirstCharacterColumnNumber               | 9  | 指定された選択範囲または指定範囲の開始位置を取得します。 選択範囲または指定範囲が解除されている場合、範囲の右側の文字番号 (ステータス バーで "桁" の後に表示される文字の列番号と同じ) を取得します。                                                              |
| wdFirstCharacterLineNumber                 | 10 | 指定された選択範囲または指定範囲の開始位置を取得します。 選択範囲または指定範囲が解除されている場合は、範囲の右側の文字番号 (ステータス バーで "行" の後に表示される文字の行番号と同じ) を取得します。                                                            |
| wdFrameIsSelected                          | 11 | 指定された選択範囲または指定範囲がレイアウト枠またはテキスト ボックス全体である場合、値は True です。                                                                                                                                                             |
| wdHeaderFooterType                         | 33 | 指定された選択範囲または指定範囲を含むヘッダーまたはフッターの種類を示す値を取得します。 詳細については、「備考」の表を参照してください。                                                                                                                         |
| wdHorizontalPositionRelativeToPage         | 5  | 指定した選択範囲または範囲の水平方向の位置を返します。これは、選択範囲または範囲の左端からページの左端までの距離です (1 ポイント = 20 twips、72 ポイント = 1 インチ)。 選択範囲または範囲が画面領域内に含めなかった場合は、-1 を返します。                        |
| wdHorizontalPositionRelativeToTextBoundary | 7  | 指定した選択範囲または範囲の水平方向の位置をポイント (1 ポイント = 20 twips、72 ポイント = 1 インチ) で、それを囲む最も近いテキスト境界の左端を基準に返します。 選択範囲または範囲が画面領域内に含めなかった場合は、-1 を返します。                               |
| wdInBibliography                           | 42 | 文献目録には、指定された選択範囲または指定範囲の場合は True を返します。                                                                                                                                                                                          |
| wdInCitation                               | 43 | 指定された選択範囲または指定範囲が引用文献の場合は True を返します。                                                                                                                                                                                              |
| wdInClipboard                              | 38 | この定数の詳細については、Microsoft Office Macintosh Edition に含まれているランゲージ リファレンスのヘルプを参照してください。                                                                                                                                    |
| wdInCommentPane                            | 26 | 指定された選択範囲または指定範囲がコメント ウィンドウ枠にある場合、値は True です。                                                                                                                                                                               |
| wdInContentControl                         | 46 | 指定された選択範囲または指定範囲がコンテンツ コントロール内にある場合は True を返します。                                                                                                                                                                         |
| wdInCoverPage                              | 41 | 送付状には、指定された選択範囲または指定範囲の場合は True を返します。                                                                                                                                                                                            |
| wdInEndnote                                | 36 | 標準表示モードで、文末脚注または文末脚注ウィンドウ枠で印刷レイアウト表示で、選択範囲または指定範囲がの場合 True を返します。                                                                                                                                      |
| wdInFieldCode                              | 44 | フィールド コードでは、指定された選択範囲または指定範囲の場合は True を返します。                                                                                                                                                                                 |
| wdInFieldResult                            | 45 | フィールドの実行結果は、指定された選択範囲または指定範囲の場合は True を返します。                                                                                                                                                                                |
| wdInFootnote                               | 35 | 標準表示モードで脚注領域または印刷レイアウト表示で、脚注ウィンドウ枠で、選択範囲または指定範囲がの場合 True を返します。                                                                                                                                          |
| wdInFootnoteEndnotePane                    | 25 | 場合、選択範囲または指定範囲の脚注または文末脚注のウィンドウで印刷レイアウト表示の脚注または文末脚注領域または標準表示モードでは、 True を返します。 詳細については、 wdInFootnote および wdInEndnote を上記の説明を参照してください。                            |
| wdInHeaderFooter                           | 28 | 場合は、選択範囲または指定範囲がヘッダーまたはフッターのウィンドウまたは、ヘッダーまたはフッターを印刷レイアウト表示では、 True を返します。                                                                                                                      |
| wdInMasterDocument                         | 34 | 選択範囲または指定範囲がグループ文書 (少なくとも 1 つのサブ文書を含む文書) 内の場合は True を返します。                                                                                                                                                           |
| wdInWordMail                               | 37 | 場合は、選択範囲または指定範囲がヘッダーまたはフッターのウィンドウまたは、ヘッダーまたはフッターを印刷レイアウト表示では、 True を返します。                                                                                                                      |
| wdMaximumNumberOfColumns                   | 18 | 選択範囲または指定範囲に含まれる表の列の最大の列数を取得します。                                                                                                                                                                                                  |
| wdMaximumNumberOfRows                      | 15 | 指定された選択範囲または指定範囲の表の最大の行数を取得します。                                                                                                                                                                                                    |
| wdNumberOfPagesInDocument                  | 4  | 選択範囲または指定範囲と関連する文書のページ数を取得します。                                                                                                                                                                                                      |
| wdNumLock                                  | 22 | Returns True if Num Lock is in effect.                                                                                                                                                                                                                            |
| wdOverType                                 | 23 | 上書きモードの場合、値は True です。 Overtype プロパティを使用して上書きモードの状態を変更できます。                                                                                                                                                              |
| wdReferenceOfType                          | 32 | 「備考」の表に示すとおり、選択範囲が脚注、文末脚注、またはコメントの参照範囲の中にあるかどうかを示す値を取得します。                                                                                                                                              |
| wdRevisionMarking                          | 24 | 変更履歴の記録がオンの場合、値は True です。                                                                                                                                                                                                                      |
| wdSelectionMode                            | 20 | 次の表に示すように、現在の選択モードを示す値を取得します。                                                                                                                                                                                                        |
| wdStartOfRangeColumnNumber                 | 16 | 選択範囲または指定範囲の先頭を含む表の列番号を取得します。                                                                                                                                                                                                        |
| wdStartOfRangeRowNumber                    | 13 | 選択範囲または指定範囲の先頭を含む表の行番号を取得します。                                                                                                                                                                                                        |
| wdVerticalPositionRelativeToPage           | 6  | 選択範囲または範囲の垂直方向の位置を返します。これは、選択範囲の上端からページの上端までの距離です (1 ポイント = 20 twips、72 ポイント = 1 インチ)。 選択範囲がドキュメント ウィンドウに表示されない場合は、-1 を返します。                                       |
| wdVerticalPositionRelativeToTextBoundary   | 8  | 選択範囲またはポイント (1 ポイント = 20 twip、72 ポイント = 1 インチ) で、それを囲む隣接する境界線の上端を基準にして範囲の垂直方向の位置を返します。 枠または表のセル内に挿入ポイントの位置を決定するのに便利です。 選択範囲が表示されない場合は、-1 を返します。 |
| wdWithInTable                              | 12 | 選択範囲が表の中にある場合、値は True です。                                                                                                                                                                                                                      |
| wdZoomPercentage                           | 19 | 割合 のプロパティが設定されている拡大率の現在の割合を返します。                                                                                                                                                                                                   |

</details>

<!-- リンク先 -->
[WdBuiltInProperty列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdbuiltinproperty
[WdRemoveDocInfoType列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdremovedocinfotype
[Document.RemovePersonalInformationプロパティ(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.document.removepersonalinformation
[WdColorIndex列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdcolorindex
[WdInformation列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdinformation

---
## 蛍光ペンの一括クリア

* Word での操作方法

    次の手順で蛍光ペンの一括クリアできる：

    1. 「検索と置換」(`Ctrl` + `H`) を表示
    2. 「オプション」を表示
    3. 「検索する文字列」の「書式」にて「蛍光ペン」を選択
    4. 「置換後の文字列」の「特殊文字列」にて「検索する文字列」を選択
    5. 「置換後の文字列」の「書式」にて「蛍光ペン」２回選択し「蛍光ペン（なし）」とする
    6. 「すべて置換」をクリック
    
* マクロの例
    <details>
    <summary>🎨 蛍光ペンの一括クリア:</summary>

    ```vb
    Sub 蛍光ペンの一括クリア()
        '' 蛍光ペンの箇所を蛍光ペン（なし）に一括置換します。
        With Application.Selection.Find
            Call .ClearFormatting
            Call .Replacement.ClearFormatting

            .Highlight = True
            .Replacement.Highlight = False

            .Text = ""
            .Replacement.Text = "^&"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False

            Call .Execute(Replace:=wdReplaceAll)
        End With
    End Sub
    ```

    </details>

## 蛍光ペンの指定色一括クリア

どうやらWordとしての機能はない模様。なのでVBAで頑張るしか方法がないかもしれない。


* マクロの例
    <details>
    <summary>🎨 蛍光ペンの指定色クリア1:</summary>

    ```vb
    Sub 蛍光ペンの指定色クリア1()
        Dim objApp As Word.Application
        Set objApp = Word.Application

        '' 蛍光ペンの箇所を検索する条件を設定します。
        With objApp.Selection.Find
            Call .ClearFormatting
            Call .Replacement.ClearFormatting

            .Highlight = True

            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
        End With

        '' 蛍光ペンの箇所をすべて検索し、指定色であればクリアします。
        Call objApp.Selection.Find.Execute(Replace:=wdReplaceNone)
        Do While objApp.Selection.Find.Found = True
            If objApp.Selection.Range.HighlightColorIndex = wdYellow Then
                objApp.Selection.Range.HighlightColorIndex = wdNoHighlight
            End If

            objApp.Selection.Collapse Direction:=wdCollapseEnd
            Call objApp.Selection.Find.Execute(Replace:=wdReplaceNone)
        Loop

        Set objApp = Nothing
    End Sub
    ```

    </details>

    <details>
    <summary>🎨 蛍光ペンの指定色クリア2:</summary>

    ```vb
    Sub 蛍光ペンの指定色クリア2()
        Dim objRange As Range
        Set objRange = ActiveDocument.Range(0, 0)

        With objRange.Find
            Call .ClearFormatting
            Call .Replacement.ClearFormatting

            .Highlight = True

            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop '' wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
        End With

        Do While objRange.Find.Execute(Replace:=wdReplaceNone) = True
            Select Case objRange.HighlightColorIndex
            Case wdYellow
                '蛍光ペン【黄色】の場合の処理
                '蛍光ペンを解除
                objRange.HighlightColorIndex = wdNoHighlight
            Case Else
                '蛍光ペン その他色 の場合の処理
            End Select

            objRange.Collapse Direction:=wdCollapseEnd
        Loop
        
        Set objRange = Nothing
    End Sub
    ```

    </details>

## 複数Wordファイルのページ番号を連番にする

* マクロの例
    <details>
    <summary>🎨 複数Wordファイルの連番化1:</summary>

    ```vb
    Sub フォルダー内文書ページ番号の連番化()

    Const USERPATH As String = "C:\@test" 'ファイルの場所
    Dim oDoc As Document
    Dim aPath As String
    Dim cnt As Long
    Dim pg As Long, pgsum As Long
    
    aPath = Dir(USERPATH & "\*.docx", vbNormal)
    Do While aPath <> ""
        cnt = cnt + 1
        Set oDoc = Documents.Open(USERPATH & "\" & aPath)
        With oDoc
        'この文書のページ数
        pg = .Content.Information(wdNumberOfPagesInDocument)
        If cnt = 1 Then
            '1番目は変更なしで閉じる
            .Close SaveChanges:=False
        Else
            '2番目以降は開始ページ番号を設定して閉じる
            .Sections(1).Footers(wdHeaderFooterPrimary) _
                .PageNumbers.StartingNumber = pgsum + 1
            .Close SaveChanges:=True
        End If
        '現在までの総ページ
        pgsum = pgsum + pg
        End With
        aPath = Dir()
    Loop
    Set oDoc = Nothing

    End Sub
    ```

    </details>

# Excel VBA マクロのメモ

## Excelカラーダイアログの表示と値の取得

* マクロの例
    <details>
    <summary>🎨 Excelカラーダイアログの表示と値の取得:</summary>

    ```vb
    Sub Excelカラーダイアログの表示と値の取得()
        Dim fullColorCode As Long
        Dim RGBRed As Integer
        Dim RGBGreen As Integer
        Dim RGBBlue As Integer

        Call Application.Dialogs(xlDialogEditColor).Show(1)
        fullColorCode = ActiveWorkbook.Colors(1)

        'Get the RGB value for each color (possible values 0 - 255)
        RGBRed = fullColorCode Mod 256
        RGBGreen = (fullColorCode \ 256) Mod 256
        RGBBlue = fullColorCode \ 65536

        Debug.Print fullColorCode; "0x" & Hex$(fullColorCode) _
            & " RGB=(" & Hex$(RGBRed) & ", " & Hex$(RGBGreen) & ", " & Hex$(RGBBlue) & ")"
    End Sub 
    ```
    
    </details>

## 塗り潰しの一括クリア

* マクロの例
    <details>
    <summary>🎨 塗り潰しの指定色クリア1:</summary>

    ```vb
    Sub 塗り潰しの指定色クリア1()
        With Application.FindFormat.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Application.ReplaceFormat.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Application.Cells.Replace _
            What:="", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False, _
            SearchFormat:=True, _
            ReplaceFormat:=True, _
            FormulaVersion:=xlReplaceFormula2
    End Sub
    ```
    
    </details>

    <details>
    <summary>🎨 塗り潰しの指定色クリア2:</summary>

    ```vb
    Sub 塗り潰しの指定色クリア2()
        Dim m_xlBook As Workbook
        Set m_xlBook = ActiveWorkbook

        With m_xlBook.Application.FindFormat.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With m_xlBook.Application.ReplaceFormat.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        m_xlBook.Application.Cells.Replace _
            What:="", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False, _
            SearchFormat:=True, _
            ReplaceFormat:=True, _
            FormulaVersion:=xlReplaceFormula2

        Set m_xlBook = Nothing
    End Sub
    ```

    </details>
