# クイックスタートガイド

## 1. セットアップ

```bash
# 依存関係をインストール
pip install openpyxl
```

## 2. サンプルファイルで試す

```bash
# テストスクリプトを実行してサンプルファイルを作成
python test_excel_row_diff.py --no-cleanup

# サンプルファイルが作成されます:
# - sample_file1.xlsx
# - sample_file2.xlsx
# - test_result.xlsx (差分結果)
```

## 3. 基本的な使い方

```bash
# 2つのExcelファイルを比較
python excel_row_diff.py file1.xlsx file2.xlsx

# 出力ファイル名を指定
python excel_row_diff.py file1.xlsx file2.xlsx -o my_result.xlsx
```

## 4. 結果の確認

出力されたExcelファイルを開くと以下のシートが含まれます：

- **summary**: 差分の概要（最初に確認するシート）
- **SheetName_only1**: ファイル1にのみ存在する行
- **SheetName_only2**: ファイル2にのみ存在する行

## 使用例シナリオ

### シナリオ1: 月次データの更新確認

```bash
# 先月と今月のデータを比較
python excel_row_diff.py data_202401.xlsx data_202402.xlsx -o changes_202402.xlsx
```

### シナリオ2: マスターデータの変更管理

```bash
# 本番環境とテスト環境のマスターデータを比較
python excel_row_diff.py master_prod.xlsx master_test.xlsx -o master_diff.xlsx
```

### シナリオ3: データ移行の検証

```bash
# 移行前後のデータを比較
python excel_row_diff.py before_migration.xlsx after_migration.xlsx -o migration_check.xlsx
```

## トラブルシューティング

### Q: "openpyxl"がインストールされていないエラー

```bash
pip install openpyxl
```

### Q: ファイルが見つからないエラー

- ファイルパスが正しいか確認してください
- 相対パスまたは絶対パスを指定できます
- ファイル名にスペースが含まれる場合は引用符で囲んでください：
  ```bash
  python excel_row_diff.py "file 1.xlsx" "file 2.xlsx"
  ```

### Q: メモリエラーが発生する

- 大きなExcelファイルの場合、メモリ不足になる可能性があります
- ファイルサイズを小さくするか、より多くのメモリを持つマシンで実行してください

## より詳しい情報

詳細な情報は [README.md](README.md) を参照してください。
