# DeliveryDocumentsCleaner
 
納品用のドキュメントを調整するツールです。

## デモ

![](assets/Animation.gif)

## 機能
指定されたフォルダ内に格納された Office ドキュメントすべてに対してハイライトと個人情報をクリアして上書き保存します。

### Word
Word 文書より以下をクリアします。
* 個人情報
* 蛍光ペン

### Excel
Excel ブックより以下をクリアします。
* 個人情報
* 塗りつぶし色

## 使用方法

1. `DeliveryDocumentsCleaner.xlam` をダウンロードし任意の場所に保存
2. `DeliveryDocumentsCleaner.xlam` を開く
   
    ![](assets/01.png)
    
    この画面が表示されたら「マクロを有効にする」をクリックしてください。

3. Excel のリボンで「アドイン」→「納品ドキュメントの調整」をクリック

    ![](assets/02.png)

    この画面が表示されたら新規ブックを作成してください。

4. 対象ドキュメントが格納されたフォルダを選択

    ![](assets/03.png)

    * 右上の「...」でフォルダを選択

5. 「蛍光ペン（Word のみ）」よりクリアする蛍光ペンをチェック
    * クリアしない場合は、すべてチェックOFFしてください
6. 「塗りつぶし色（Excel のみ）」よりクリアする塗りつぶし色を選択
    * 選択した色を解除する場合は「色の設定」で「キャンセル」してください
7. 「各クリア実行」をクリック
    * ⚠️ 対象ドキュメントは事前にすべて閉じてください
    * ⚠️ 対象ドキュメントは上書き保存します