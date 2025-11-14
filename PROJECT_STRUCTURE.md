# プロジェクト構成

```
excel-row-diff/
├── README.md                # メインドキュメント
├── QUICKSTART.md           # クイックスタートガイド
├── LICENSE                 # MITライセンス
├── pyproject.toml         # Pythonプロジェクト設定
├── .gitignore             # Git除外設定
├── excel_row_diff.py      # メインスクリプト
└── test_excel_row_diff.py # テストスクリプト
```

## ファイルの説明

### excel_row_diff.py
メインの実行ファイル。2つのExcelファイルを比較し、差分を抽出します。

**主な機能:**
- Excelファイルの読み込み
- シートの自動検出
- 行単位の比較
- 差分の抽出と出力
- サマリーレポートの生成

### test_excel_row_diff.py
テスト用スクリプト。サンプルExcelファイルを自動生成して動作確認を行います。

**実行方法:**
```bash
python test_excel_row_diff.py
```

### pyproject.toml
Pythonプロジェクトの設定ファイル。依存関係やパッケージメタデータを定義します。

### .gitignore
Gitで管理しないファイルを指定。テストファイルやビルド成果物などを除外します。

## 使用時に生成されるファイル

実行すると以下のファイルが生成されます：

```
diff_result.xlsx        # デフォルトの出力ファイル名
verification_output.xlsx # -o オプションで指定した場合
sample_file1.xlsx       # テストスクリプトで生成
sample_file2.xlsx       # テストスクリプトで生成
test_result.xlsx        # テストスクリプトの出力
```

これらのファイルは `.gitignore` で除外されているため、Gitにコミットされません。
