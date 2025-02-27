# AutoDiary_v2-LINEbotGPT-o3-mini

LINEボットを通じて自動的に日記を作成・管理するシステムです。ChatGPT（o3-mini）を活用して、ユーザーとの自然な会話からパーソナライズされた日記エントリを生成します。

## 概要

このプロジェクトは、LINEのMessaging APIとOpenAIのChatGPT APIを組み合わせて、ユーザーとの会話を通じて日記を作成するボットを実装しています。Google Apps Script上で動作し、スプレッドシートをデータベースとして使用しています。

## 主な機能

- **LINEボットとの会話**: ユーザーはLINEアプリを通じてボットと会話できます
- **ChatGPTによる応答生成**: ユーザーのメッセージに対して、ChatGPT（o3-mini）を使用して自然な応答を生成
- **会話履歴の管理**: ユーザーごとに会話履歴を保存し、コンテキストを維持
- **複数ユーザー対応**: 複数のユーザーが同時にアクセスしても適切に処理
- **非同期処理**: メッセージキューシステムによる効率的なリクエスト処理
- **エラー処理**: 通信エラー時の自動リトライ機能

## 技術的な特徴

- **非同期処理アーキテクチャ**: キューシステムを使用して複数リクエストを効率的に処理
- **バッチ処理**: 複数のChatGPTリクエストを一括処理
- **ユーザー管理**: 各ユーザーの情報と会話履歴を個別に管理
- **エラーハンドリング**: 指数バックオフによるリトライ機能

## セットアップ方法

1. **Google Apps Scriptプロジェクトの作成**
   - Google Driveで新しいスプレッドシートを作成
   - ツール > スクリプトエディタを選択
   - コードをコピーして貼り付け

2. **スプレッドシートの設定**
   - 以下のシートを作成:
     - プロンプト: ChatGPTへの指示を設定
     - 検証: Webhook接続確認用
     - ログ: 全体の会話ログ
     - ユーザー: ユーザー情報管理
     - 情報: 日記データ

3. **API設定**
   - LINE Developers ConsoleでMessaging APIチャネルを作成
   - OpenAIでAPIキーを取得
   - スクリプトプロパティに以下を設定:
     - `apikey`: OpenAI APIキー
     - `linetoken`: LINE Messaging APIトークン

4. **デプロイ**
   - ウェブアプリとしてデプロイ
   - 取得したURLをLINE DevelopersのWebhook URLに設定

5. **トリガー設定**
   - `setupTriggers()`関数を実行してキュー処理用のトリガーを設定

## 使用方法

1. LINEアプリでボットを友達追加
2. メッセージを送信して会話を開始
3. ボットはChatGPTを使用して応答し、会話履歴を保存

## 参照スプレッドシート

[AutoDiary_v2 スプレッドシート](https://docs.google.com/spreadsheets/d/17idcpzqeSxOf0apZZYBm76C-lNuf4qVjwfBn5tGaU7o/edit?usp=sharing)

## 注意事項

- ChatGPT APIの使用には料金が発生します
- LINE Messaging APIの無料枠には制限があります
- スプレッドシートの容量には上限があるため、長期運用時は定期的なメンテナンスが必要です

## 技術スタック

- Google Apps Script
- LINE Messaging API
- OpenAI ChatGPT API (o3-mini)
- Google Spreadsheet
