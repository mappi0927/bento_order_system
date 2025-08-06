# 弁当注文管理システム（Bento Order Management System）

## 📌 概要

このソフトは、会社内の昼食（お弁当）の注文業務を効率化するために作成された Python ベースの業務支援ツールです。  
紙の注文票と手作業での集計から、CSV・FAX注文・月次精算までを一元管理できます。

## 🛠 主な機能

- 商品コード＆個数をCSVに記入して、注文一覧表を自動生成
- 弁当業者ごとのFAX注文票（Excelテンプレート）を作成
- 注文データを累積保存（履歴管理）
- 日付指定での注文集計（社員別・部門別・業者別）
- 月額精算：会社負担額と給与天引き額を自動計算（仕様をご指示ください）
- CSVプレビュー→印刷支援用のVBS/BATスクリプトも同梱

## 📁 フォルダ構成
bento_order_system/
├── main.py                  # メイン処理ファイル
├── utils.py                 # 補助関数
├── CSV/
│   ├── uriage010_nemu.csv       # メニューマスター
│   ├── uriage020_hibi.csv       # 社員マスター（注文入力）
│   ├── uriage030_ruiseki.csv    # 注文履歴（累積）
│   ├── FAX_Order_Summary_1.csv  # FAX用注文CSV（業者11）
│   ├── FAX_Order_Summary_2.csv  # FAX用注文CSV（業者22）
├── Excel_Template/
│   ├── FAX_Order_Summary_11.xlsx  # FAX注文票（業者11用）
│   ├── FAX_Order_Summary_22.xlsx  # FAX注文票（業者22用）
├── Scripts/
│   ├── 010_hiraku3.vbe        # CSVを開くVBSスクリプト
│   ├── 020_tojiru.vbe         # CSVを閉じるVBSスクリプト
```

## 🔧 使用方法（概要）

1. `uriage020_hibi.csv` に注文データ（コード・個数）を入力  
2. Pythonスクリプトを実行して、当日注文一覧やFAX注文CSVを生成  
3. CSVはExcelテンプレートと連携し、プレビュー後に印刷  
4. 日次の注文は累積記録され、月末に集計・精算処理を実施

## 💡 想定利用者

- 中小企業で日々の昼食注文を手作業で集計している事務担当者
- 業務改善や業務自動化を検討している社内SEや現場リーダー

## 📌 補足

- 実在の社員名・業者名・価格情報は含まれていません（ダミーデータ）
- 印刷補助用のVBS/BATスクリプトは、Windows環境での利用を想定しています

---

👤 作者：mappi0927  
🕒 公開日：2025年8月  

