画像ファイルをエクセルに自動で貼り付けるスクリプトです。

### 設定
config.jsonに記述します。
- dpi: モニターのDPI
- max_height: 画像リサイズ時の最大高さ（ポイント）
- max_width: 画像リサイズ時の最大幅（ポイント）
- n_cols：画像を何枚横に並べるか
- col_gap：横並びの画像の間隔（シートの列数）
- row_gap_small：同じ項番内で改行する際に空ける幅（シートの行数）
- row_gap_big：項番が切り替わるときに空ける幅（シートの行数）
- row_height_point: シートの一行が何ポイントか。リサイズに使用。
- prefix: 項番の接頭辞。
- suffix: 項番の接尾辞。

### 実行
実行の前提：
- あるディレクトリの配下に1, 2, 3など整数値の名前を持つ子ディレクトリがある。連番でなくてもいい。
- ディレクトリ1ならば、1-1.png, 1-2.pngの形式の名前のpngを含む。連番でなくてもいい。

```bash
python paster.py --dirname testdir --out test.xlsx --prefix 第 --suffix 項
```

この場合testdir配下に1, 2, 3などpngが格納されたディレクトリがあります。
出力されるのはpngが貼り付けられたtest.xlsxです。
各子ディレクトリの前に「第1項」などのラベルが書き込まれます。

オプションは以下の通りです。
- --dirname / -d : 画像保存先の親ディレクトリ。必須。
- --out / -o : 出力するブック名。必須。拡張子.xlsxをつけなければ補われる。
- --prefix / -p：項番の接頭辞。省略可能で設定ファイルより優先される。
- --suffix / -s：項番の接尾辞。省略可能で設定ファイルより優先される。
- --help / -h：ヘルプを表示する。