# 自動タスク生成 全手順ガイド（最終版）

-----

## 全体構成

```
SharePoint /Shared Documents/
    ├─ 元データ.xlsx        ← 書き込み禁止・日々追記される
    ├─ 管理.xlsx            ← 最終読込行を保存
    └─ 作業用/
         └─ 作業用データ.xlsx ← Power Automateが毎回コピー作成
```

```
毎日9:00 Power Automate起動
        │
        ▼
① 管理.xlsxから最終読込行を取得
        │
        ▼
② 元データ.xlsx → 作業用データ.xlsxにコピー
        │
        ▼
③ Office Script「フラグ付け」実行
  （最終読込行+1から新規行のみ処理）
  A列「リモート」またはB列黄色 → D列に○
        │
        ▼
④ 管理.xlsxの最終読込行を更新
        │
        ▼
⑤ Power Query更新
  ○の行のみ抽出 → グループ化 → タスク一覧完成
```

-----

## STEP 1：Excelファイルの準備

### 1-1. 管理.xlsxを新規作成

```
SharePointの /Shared Documents/ に「管理.xlsx」を新規作成

Sheet1の構成：
┌───────────────┬───┐
│ 最終読込行    │ 0 │  ← B1に行番号を保存（初期値は0）
└───────────────┴───┘
```

### 1-2. 元データ.xlsxの列構成を確認

```
┌──────┬──────┬──────┬──────┐
│ A列  │ B列  │ C列  │ D列  │
├──────┼──────┼──────┼──────┤
│担当者│機器  │依頼No│フラグ│ ← 1行目：ヘッダー
├──────┼──────┼──────┼──────┤
│Aさん │A機   │ 122  │      │ ← 2行目以降：データ
│Bさん │B機   │ 122  │      │
│Zさん │A機   │ 123  │      │
└──────┴──────┴──────┴──────┘
※ D列（フラグ）がOffice Scriptで○を入れる列
※ データに空行は入らない運用とする
```

### 1-3. 作業用フォルダを作成

```
SharePoint /Shared Documents/作業用/
※ フォルダだけ作成しておく（中身は空でOK）
```

-----

## STEP 2：黄色の色コード確認

```
元データ.xlsx をブラウザで開く
↓
「自動化」タブ →「新しいスクリプト」
↓
以下を貼り付けて実行
```

```typescript
// 色コード確認用（確認後は削除してOK）
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  // 黄色のセルがある行・列に合わせて変更
  const color = sheet.getCell(1, 1).getFormat().getFill().getColor();
  console.log(color); // 例：#FFFF00
}
```

```
コンソールに出力された値をメモしておく
（STEP3のYELLOWに設定する）
```

-----

## STEP 3：Office Scriptを4つ作成

### 作成場所

```
任意のExcelファイルをブラウザで開く
↓
「自動化」タブ →「新しいスクリプト」
↓
スクリプト名と内容を設定して保存
```

-----

### Script① 「最終行取得」

#### 実行対象：管理.xlsx

```typescript
function main(workbook: ExcelScript.Workbook): number {
  const sheet = workbook.getActiveWorksheet();
  const lastRow = Number(sheet.getCell(0, 1).getValue()); // B1の値
  return isNaN(lastRow) ? 0 : lastRow;
}
```

-----

### Script② 「フラグ付け」

#### 実行対象：作業用データ.xlsx

```typescript
function main(workbook: ExcelScript.Workbook, lastReadRow: number): number {

  const sheet = workbook.getActiveWorksheet();
  const YELLOW = "#FFFF00"; // STEP2で確認した色コードに変更

  // 開始行を決定（初回は1行目=ヘッダーの次、2回目以降は最終行+1）
  const startRow = (lastReadRow === 0 || isNaN(lastReadRow)) ? 1 : lastReadRow;

  // データの最終行を取得（空行なし運用のためgetUsedRangeで正確に取得可能）
  const totalRows = sheet.getUsedRange().getRowCount();

  // 新規行がなければスキップ
  if (startRow >= totalRows) {
    console.log("新規行なし、処理をスキップします");
    return totalRows;
  }

  console.log(`処理開始行: ${startRow + 1}行目 ～ ${totalRows}行目`);

  // 新規行のみループ処理
  for (let r = startRow; r < totalRows; r++) {
    const cellA = sheet.getCell(r, 0); // A列：担当者
    const cellB = sheet.getCell(r, 1); // B列：機器
    const cellD = sheet.getCell(r, 3); // D列：フラグ

    const valueA = String(cellA.getValue());
    const colorB = cellB.getFormat().getFill().getColor().toUpperCase();

    if (valueA.includes("リモート") || colorB === YELLOW) {
      cellD.setValue("○");
    }
  }

  // 処理した最終行番号をPower Automateへ返す
  return totalRows;
}
```

-----

### Script③ 「最終行保存」

#### 実行対象：管理.xlsx

```typescript
function main(workbook: ExcelScript.Workbook, newLastRow: number): void {
  const sheet = workbook.getActiveWorksheet();
  sheet.getCell(0, 1).setValue(newLastRow); // B1に上書き
  console.log(`最終読込行を ${newLastRow} に更新しました`);
}
```

-----

### Script④ 「Query更新」

#### 実行対象：作業用データ.xlsx

```typescript
function main(workbook: ExcelScript.Workbook): void {
  workbook.refreshAllDataConnections();
  console.log("Power Query更新完了");
}
```

-----

## STEP 4：Power Queryを作成

### 4-1. 作業用データ.xlsxにPower Queryを設定

```
作業用データ.xlsx をブラウザで開く
↓
「データ」タブ →「データの取得」→「ブックから」
→ 作業用データ.xlsx自身のSheet1を指定
↓
Power Query エディター →「詳細エディター」
↓
以下を貼り付け
```

```m
let
    // ① データ取得
    Source = Excel.CurrentWorkbook(){[Name="Sheet1"]}[Content],

    // ② 1行目をヘッダーに昇格
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),

    // ③ D列（フラグ）が○の行のみ抽出
    FilteredRows = Table.SelectRows(
        Headers,
        each [フラグ] = "○"
    ),

    // ④ 機器列をアンピボット（複数機器を縦展開）
    Unpivoted = Table.UnpivotOtherColumns(
        FilteredRows,
        {"担当者", "依頼No"},
        "属性",
        "機器"
    ),

    // ⑤ 属性列を削除
    Cleaned = Table.RemoveColumns(Unpivoted, {"属性"}),

    // ⑥ 依頼No・機器でグループ化 → 担当者を「、」で結合
    Grouped = Table.Group(
        Cleaned,
        {"依頼No", "機器"},
        {{
            "担当者",
            each Text.Combine(List.Sort([担当者]), "、"),
            type text
        }}
    ),

    // ⑦ ソート
    Sorted = Table.Sort(
        Grouped,
        {{"依頼No", Order.Ascending}, {"機器", Order.Ascending}}
    ),

    // ⑧ タスク番号付与
    WithIndex = Table.AddIndexColumn(Sorted, "タスクNo", 1, 1, Int64.Type),

    // ⑨ 列の並び替え
    Result = Table.ReorderColumns(
        WithIndex,
        {"タスクNo", "機器", "担当者", "依頼No"}
    )
in
    Result
```

```
「閉じて読み込む」で完了
```

-----

## STEP 5：Power Automateフローを作成

### 5-1. フロー新規作成

```
https://make.powerautomate.com
↓
「作成」→「スケジュール済みクラウドフロー」
・フロー名：毎日タスク自動生成
・繰り返し：1日（毎日9:00）
```

### 5-2. アクションを順番に追加

```
① トリガー：スケジュール（毎日9:00）
        │
        ▼
② Excel Online「スクリプトの実行」
   ・ファイル  ：管理.xlsx
   ・スクリプト：最終行取得
        │
        ▼
③ 変数の初期化
   ・名前：lastReadRow
   ・型  ：整数
   ・値  ：②のresult（動的コンテンツ）
        │
        ▼
④ SharePoint「ファイルのコピー」
   ・コピー元  ：/Shared Documents/元データ.xlsx
   ・コピー先  ：/Shared Documents/作業用/
   ・新しい名前：作業用データ.xlsx
   ・上書き    ：はい
        │
        ▼
⑤ Excel Online「スクリプトの実行」
   ・ファイル         ：④のID（動的コンテンツ）
   ・スクリプト       ：フラグ付け
   ・引数 lastReadRow ：③の変数
        │
        ▼
⑥ 変数の更新
   ・lastReadRow = ⑤のresult（動的コンテンツ）
        │
        ▼
⑦ Excel Online「スクリプトの実行」
   ・ファイル        ：管理.xlsx
   ・スクリプト      ：最終行保存
   ・引数 newLastRow ：⑥の変数
        │
        ▼
⑧ Excel Online「スクリプトの実行」
   ・ファイル  ：④のID（動的コンテンツ）
   ・スクリプト：Query更新
```

-----

## 最終チェックリスト

```
【STEP1 ファイル準備】
□ 管理.xlsxを作成・SharePointに配置
□ 管理.xlsxのB1の初期値を 0 に設定
□ 元データ.xlsxの列構成（A〜D列）を確認
□ 作業用フォルダを作成

【STEP2 色コード確認】
□ 黄色の色コードを確認してメモ

【STEP3 Office Script】
□ Script①「最終行取得」を作成・保存
□ Script②「フラグ付け」を作成・YELLOWの値を修正して保存
□ Script③「最終行保存」を作成・保存
□ Script④「Query更新」を作成・保存

【STEP4 Power Query】
□ Power Queryを作成・動作確認

【STEP5 Power Automate】
□ フローを作成
□ 手動実行でテスト・結果を確認
□ スケジュール実行を有効化
```