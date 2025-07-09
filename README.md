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


## F#でExcelを操作するコード例

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