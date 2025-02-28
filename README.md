# IIIF対応画像から歴史学研究に活用するRDFデータを作成するためのワークフロー
本リポジトリは、18世紀フランスの議事録の目録をRDFデータ化する際に使用したコードやツールについての情報をまとめています。
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Documentation: CC BY 4.0](https://img.shields.io/badge/Docs-CC%20BY%204.0-blue.svg)](LICENSE)
[![Last Updated](https://img.shields.io/github/last-commit/ayanosk/historical-rdf-workflow)](https://github.com/ayanosk/historical-rdf-workflow)

## 目次
- [特徴](#特徴)
- [作業フローの概要](#作業フローの概要)
- [表形式データの作成](#表形式データの作成)
- [RDF（Turtle形式）への変換](#RDF（Turtle形式）への変換)
- [ライセンス](#ライセンス)
- [問い合わせ](#問い合わせ)

## 特徴
- 既存の無料ツールを組み合わせたデータ作成
- 矩形ではなく点によるIIIF画像とテキストの紐付け
- 重複項目の自動入力
- 複数人による同時作業

## 必要な環境
- 良好なウェブ接続のみ（接続状況が悪いと自動入力が機能しないため）

## 作業フローの概要
1. IIIFマニフェストを取得し、Recogitoで読み込む
2. GoogleスプレッドシートをGoogle Apps Scriptを用いて編集（preprocessing.gs）
3. Recogitoにアノテーションをつけながら複数人でスプレッドシート編集
4. RDFのプロパティを決定してGoogleスプレッドシートのヘッダーを編集する
5. OpenRefineでTurtle形式に変換する
6. GraphDBで読み込み、分析する

## 表形式データの作成

## RDF（Turtle形式）への変換

## ライセンス
- コード: MIT License - https://opensource.org/licenses/MIT
- ドキュメント: CC BY 4.0 - https://creativecommons.org/licenses/by/4.0/deed.ja

## 問い合わせ
バグ報告や提案がある場合は、Issue を作成してください。

- 開発者:小風綾乃（researchmap: https://researchmap.jp/ayano_sanno)



