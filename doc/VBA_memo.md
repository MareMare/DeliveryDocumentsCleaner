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

<!-- リンク先 -->
[WdBuiltInProperty列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdbuiltinproperty
[WdRemoveDocInfoType列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdremovedocinfotype
[Document.RemovePersonalInformationプロパティ(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.document.removepersonalinformation
[WdColorIndex列挙(Word)]: https://learn.microsoft.com/ja-jp/office/vba/api/word.wdcolorindex

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
