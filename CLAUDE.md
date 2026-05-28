# CLAUDE.md — webmark_mania_news プロジェクト

## プロジェクト概要

マニアシリーズ6店舗のお知らせ内容を管理するプロジェクト。
Instagramからピックアップした投稿をサイトへ転載する際の被り防止チェックが主な用途。

## 対象店舗

- 小籠包マニア 中目黒本店
- 焼き小籠包マニア
- 北京ダックマニア 虎ノ門ヒルズ本店
- 餃子マニア 虎ノ門ヒルズ店
- 餃子マニア 品川別館
- 餃子マニア 品川本店

## フォルダ構成

```
webmark_mania_news/
├── news_list/
│   ├── get_news_list.py              # 各店舗サイトからお知らせを収集するスクリプト
│   └── YYYY年MM月_お知らせ内容リスト(2年間).xlsx  # 照合の基準ファイル
└── .claude/
    ├── rules/news-checker.md         # 照合ルール
    └── skills/news-checker/SKILL.md  # 被りチェックスキル
```

## 主な使い方

### 月次作業
1. `news_list/get_news_list.py` を実行してExcelを最新化する
2. ファイル名を `YYYY年MM月_お知らせ内容リスト(1年間).xlsx` の形式で保存する

### 被りチェック
InstagramのURLまたは投稿内容の説明を送るだけで照合できる。
詳細は `.claude/skills/news-checker/SKILL.md` を参照。

## ルール参照

- 被りチェック: `.claude/rules/news-checker.md`
