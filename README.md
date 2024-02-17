## Excel to JSON コンバーターマニュアル

### 機能:
1. Excel ファイルを JSON 形式に変換します。
2. 変換された JSON ファイルは、元の Excel ファイルと同じディレクトリに保存されます。
3. コマンドラインから実行される際には、対話的なオプションが提供されます。

### 使用方法:

- コマンドラインからプログラムを実行する場合:
  - `python プログラム名.py`: カレントディレクトリ内の Excel ファイルを処理します。
  - `python プログラム名.py ファイルまたはフォルダのパス`: 指定された Excel ファイルまたはフォルダ内の Excel ファイルを処理します。

- Excel ファイルが処理される際には、指定された Excel ファイルの内容に基づいて適切な JSON ファイルが生成されます。

- コマンドラインからプログラムを実行する場合、処理を実行するかどうかの確認が求められます。

### コンポーネント:

1. **ConvertExcelDir(directory)**:
   - 機能: 指定されたディレクトリ内の Excel ファイルを検索し、それらを JSON に変換します。
   - 引数: `directory` - Excel ファイルが検索されるディレクトリのパス。
   - 返り値: なし。

2. **IsQripExcelFormat(active_sheet)**:
   - 機能: アクティブな Excel シートが特定の形式に適合しているかどうかを判定します。
   - 引数: `active_sheet` - アクティブな Excel シート。
   - 返り値: True（適合している場合）または False（適合していない場合）。

3. **Sheet2Json(active_sheet, excel_file_path)**:
   - 機能: Excel シートを JSON 形式に変換します。
   - 引数: `active_sheet` - アクティブな Excel シート、`excel_file_path` - 元の Excel ファイルのパス。
   - 返り値: なし。

4. **Excel2Json(excel_file_path)**:
   - 機能: Excel ファイル全体を処理し、適合するシートを JSON に変換します。
   - 引数: `excel_file_path` - 変換する Excel ファイルのパス。
   - 返り値: なし。

5. **main()**:
   - 機能: プログラムのエントリーポイント。コマンドライン引数を処理し、処理対象の Excel ファイルを指定します。
   - 引数: なし。
   - 返り値: なし。
