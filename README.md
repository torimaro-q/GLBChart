# GLBChart

Excel VBAからGLB形式の3Dチャートらしきものをつくるやつです

![pic1](pic/screen2.gif)

## Overview / 概要

GLBChart は、Excelのセル範囲にある数値データを GLB（glTF Binary） 形式の3Dモデルに変換し、そのままExcelのワークシートへ3Dモデルとして挿入するVBAクラスです

## Features / 特徴
- 外部3Dライブラリ不要
- ExcelのXYZ数値データを3Dモデルへ変換
- 単一ファイルで完結するGLBを生成
- 生成したGLBをExcelへ直接挿入
- マーカー・ラインの色を変更可能
- データ点を細い直方体で接続して表示
- X・Y・Z軸も直方体で生成
- XY平面に透明な床を直方体で生成
- だいたい全部直方体
- だいたい全部VBA

## Requirements / 動作環境

Shapes.Add3DModel が利用できるExcel環境が必要です


## Installation / 導入

GLBChartBuilder をインポートして完了です
想定する入力形式は以下を参照してください
最初に現れる「数値ではない行」は、軸ラベルとして扱います

|Temperature|Pressure|Altitude|
|:---:|:---:|:---:|
|20|101|10|
|25|98|20|
|30|95|30|

**軸ラベル**
- X = Temperature
- Y = Pressure
- Z = Altitude

現在の実装では、これらのラベルはチャート定義のために内部的に使用されます。

## Basic Usage / 基本的な使い方

```vb
Public Sub sample1()
    With New GLBChartBuilder
        .HasLine = True
        .HasMarker = False
        .LineWidth = 5
        .LineColor = RGB(50, 150, 50)
        With .CreateFromRange(ActiveSheet.Range("B1:D5000"))
            .Left = 30
            .Top = 25
            .width = 250
            .Height = 250
        End With
    End With
End Sub
```

指定した範囲から3Dモデルを生成します。
戻り値はExcelの Shape オブジェクトです。

なお、外観を変更するためのプロパティがあります。
### Markers

(例)
- HasMarker = True
- MarkerSize = 10
- MarkerColor = RGB(50, 100, 200)

データ点は小さな立方体(直方体)として表現されます。

### Lines
(例)
- HasLine = True
- LineWidth = 1
- LineColor = RGB(30, 30, 30)

有効にすると、連続するデータ点が矩形断面の直方体で接続されます。

**やはり直方体、直方体はすべてを解決する。**


## 仕様
生成されるモデルには以下が含まれます。

- データ点
- 接続線
- X軸
- Y軸
- Z軸
- 半透明の床

入力データは正規化され、おおむね1×1×1の座標空間に収まるように変換されます。

そのため、生成された3Dモデルの座標値は、元データの数値そのものではありません。

**これは仕様です。**

## 仕組み

GLBChartは外部のGLTFライブラリに依存していません。
VBAから直接GLBファイルを組み立てます。

基本的な処理の流れは以下です。
```
Excel Range
    ↓
Vector3 data
    ↓
Geometry generation
    ↓
Binary vertex/index buffers
    ↓
glTF JSON
    ↓
GLB serialization
    ↓
Temporary .glb file
    ↓
Excel 3D Model
```


実装では以下を手動で生成します。
```
glTF materials
buffer views
accessors
meshes
primitives
vertex buffers
index buffers
GLB header
JSON chunk
BIN chunk
```

**VBAすごい**

## Limitations / 制限事項

GLBChartは意図的に小さく、シンプルに作られています。

現在の主な制限事項：
- ライティング設定UIなし
- テクスチャなし
- 法線なし
- アニメーションなし
- 3Dモデル内への文字ラベル表示なし
- 自動衝突判定なし
- データ数が多いとGLBが巨大になる
- マーカーはGPUインスタンシングではなく個別のジオメトリとして生成


## License / ライセンス
MIT License

## Status / ステータス
Experimental / 実験的

GLBChartは、Excel VBAから直接シンプルな3D可視化を生成するための、小さな実験的プロジェクトです。

# Disclaimer / 免責事項

本プロジェクトは現状のまま提供されます。
自己責任で可視化してください。
