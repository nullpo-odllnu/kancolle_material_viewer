# kancolle_material_viewer

## なにこれ
- 艦これの資源をexcel管理したものを可視化するためのスクリプト
- excelのグラフでいいけど、より詳細なものが欲しい
- excelファイルを入力に、以下のような可視化htmlを出力する
    - <img src = "viewer.gif">
- 時間単位での定期実行も実装

## 使用例
- `python kancolle_material_viewer.py -e <入力excelファイルパス> -f 12 -c 備考`
    - 実行ディレクトリに `kancolle_material_viewer.html` が出力される
    - 12時間おきに出力ファイル更新
    - 備考列をコメントとして扱う

## excel仕様
- 以下のシートがあることを想定して実装している
    - `資源メモ` (コマンドライン引数で変更可能)
- 以下のカラムがあることを想定して実装している
    - 燃料
    - 弾薬
    - 鉄鋼
    - ボーキ
    - バケツ
    - (任意の列)
        - コマンドライン引数で特定の列をコメントとして使用する
        - デフォルトでは未使用
- 詳細はリポジトリ内のテスト用excelファイルを参照
    - バケツのデータが最初無いのは、作者が途中で記録し始めたため
