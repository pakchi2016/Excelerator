# Excelerator
 - ExcelをF#で操作するためのプロジェクト

## プロジェクトフォルダ作成
 - 以下コマンドを実行し、プロジェクトフォルダを作成する
```
dotnet new console -lang F# -o Excelerator
```


### ライブラリ組み込み
 - Exceleratorフォルダ内に[nuget.config]というファイルを作成し、以下の内容を記述する
```
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="Local Packages" value="C:work\開発\FShrp\Library" />
  </packageSources>
</configuration>
```

 - 上記が準備出来たら以下コマンドを実行し、Excelerator.vbprojに定義を組み込む
```
dotnet add package ClosedXML
```

## F#でExcelを操作するコード例
 - Exceleratorフォルダ配下の以下ファイルにコードを記述

```
Program.fs
```
 - 次に、上記ファイルに対して以下コードを入力

```
open ClosedXML.Excel
open System.Linq

// --- 1. 設定 ---
let sourceFile = "C:/path/to/your/ブックA.xlsx"
let destFile = "C:/path/to/your/ブックB.xlsx"

printfn "処理を開始しますわ..."

// --- 2. メイン処理 ---
// 両方のブックを最初に開いておきます
use sourceWb = new XLWorkbook(sourceFile)
use destWb = new XLWorkbook(destFile)

let sourceWs = sourceWb.Worksheet("シートA")
let destWs = destWb.Worksheet("シートX")
let startRow = 27

// --- 真のストリーミングパイプライン ---
sourceWs.RowsUsed()
// |> シートAで使われている全ての行を、一つずつ流し始めます

|> Seq.filter (fun row ->
    // |> 各行をその場で判定します
    match row.Cell(12).Value.TryGetText() with
    | true, "〇" -> true
    | _ -> false
)
// |> Seq.toList を削除！ これでメモリに溜め込むことはありません

|> Seq.iteri (fun index sourceRow ->
    // |> 条件を満たした行が一つ見つかるたびに、この処理が即座に実行されます
    if index = 0 then
        printfn "条件に合う行が見つかりました。書き込みを開始します..."

    let destRowNumber = startRow + index
    
    sourceRow.Cells()
    |> Seq.iter (fun sourceCell ->
        destWs.Cell(destRowNumber, sourceCell.Address.ColumnNumber).Value <- sourceCell.Value
    )
)

// 全ての処理が終わった後で、一度だけ保存します
destWb.Save()
```

 - コード入力完了後、以下コマンドでコードを実行

```
dotnet run
```

 - また、毎回必ず同じコードを実行する場合、以下コマンドでスタティックリンク化しておくのもよき

```
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```


### おまけ
 - C#でカレントを取得するコード

```
string exeDirectory = System.AppContext.BaseDirectory;
string settingFilePath = System.IO.Path.Combine(exeDirectory, "setting.ini");
```