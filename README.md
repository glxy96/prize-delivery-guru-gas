# ガチャ特典配布管理ツール (prize-delivery-guru)

## 概要

本ツールは、ガチャツールの結果に基づいて、デジタルグッズ特典をリスナーごとに自動的にフォルダ分けし、配布するためのGoogle Apps Scriptプロジェクトです。スプレッドシートとGoogle Driveを活用して、ガチャの結果を解析し、各リスナーに適切な特典ファイルを整理・配布する作業を効率化します。

## 主な機能

- ガチャ結果テキストの解析
  - chrome拡張機能「なまずガチャ履歴吐き出し」を利用します。
  [なまずガチャ出力](https://nonkotobuki.dokkoisho.com/namazu_gacha_syukei.html)
- リスナーごとの特典リスト作成
- リスナー別の特典フォルダの自動作成
  - 特典ファイルの自動コピー
  - 大量データでもタイムアウトしないバッチ処理
- 詳細なデバッグ情報の提供

## セットアップ方法

1. Google スプレッドシートを新規作成します
2. 拡張機能 > Apps Scriptをクリックし、スクリプトエディタを開きます
3. プロジェクトに以下のファイルを追加します：
   - `code.gs`
   - `aggregate.gs`
   - `distributePrize.gs`
4. スクリプトをコピー＆ペーストしてください
5. スプレッドシートに戻り、ページを更新します
6. 「なまずガチャ特典配布管理」メニューが表示されます
7. 「0:初期設定」を実行します

## 使用方法

### 初期設定

1. メニューの「なまずガチャ特典配布管理」>「0:初期設定」を実行します
2. Google Drive上に「prize-guru」フォルダと「prizes」サブフォルダが作成されます。Googleドライブのマイドライブ直下を確認してください。
3. 「prizes」フォルダに特典ファイルをアップロードします（ファイル名はガチャの「景品名」と**完全に一致**させてください）
  例:景品名に`アイコンリング1`を指定した場合、`prize-guru/prizes/アイコンリング1.png`を用意してください

### ガチャ結果の解析

1. 「ガチャ結果入力」シートを開きます
2. スプレッドシートのメニューから「ファイル」>「インポート」を選択します
3. 「アップロード」タブで「なまずガチャ履歴吐き出し」で取得したテキストファイルを選択します
4. インポート設定で以下を選択します：
   - 「既存のシートの内容を置き換える」
   - 区切り文字：「タブ」
5. 「インポート」ボタンをクリックします
6. メニューの「なまずガチャ特典配布管理」>「1:ガチャ結果を解析」をクリックします
7. 「配布リスト」シートが作成され、リスナーごとの特典が表示されます

### 特典ファイルのフォルダ化

1. メニューの「なまずガチャ特典配布管理」>「2:特典ファイルをフォルダ化」をクリックします
2. 確認ダイアログで「はい」をクリックします
3. 処理が始まり、各リスナーごとの特典フォルダが作成されます
4. 処理状況は「デバッグ情報」シートで確認できます
5. 処理が完了すると、リスナーごとの共有URLが配布リストに入力されます

## 注意事項

- 特典ファイルの名前は、ガチャの「景品名」と一致させてください（例：「アイコンリング1.jpg」）
- 特典ファイルは必ず「prize-guru/prizes」フォルダに置いてください
- 処理中はスプレッドシートを開いたままにしておいてください
- 大量のデータを処理する場合は、自動的に分割処理されます
- 処理をキャンセルしたい場合は「フォルダ化処理をキャンセル」を実行してください

## トラブルシューティング

- 特典ファイルが見つからない場合は、ファイル名とガチャの景品名が完全に一致しているか確認してください
- 日本語のファイル名で問題が発生する場合は、ファイル名にユニコード正規化が適用されています
- 詳細なエラー情報は「デバッグ情報」シートで確認できます

## 技術情報

- Google Apps Script を使用して開発されています
- データ処理の時間制限に対応するためのバッチ処理機能を実装しています
- 複数のガチャ特典を持つリスナーには別シートで詳細情報を提供します
- ファイル名のユニコード正規化処理によって日本語ファイル名の互換性を確保しています

## 開発者向け情報

- `code.gs` - メインコードとセットアップ機能
- `aggregate.gs` - ガチャ結果解析・集計機能
- `distributePrize.gs` - 特典フォルダ作成とバッチ処理機能
