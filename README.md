# マニアプロデュース様 お知らせ管理

マニアシリーズ6店舗のサイトお知らせを管理するプロジェクトです。
Instagramや各種媒体からピックアップした投稿をサイトへ転載する際の被り防止チェックと記事生成を自動化します。

---

## 使い方（Claude Code スキル）

**Claude Code でこのフォルダを開いた状態で**、以下の形式で送るだけで被りチェック→記事生成→デスクトップ保存まで自動で行います。

```
店舗名：〇〇マニア 〇〇店
お知らせ素材：
（テキスト貼り付け、または URL、または内容の説明）
```

### 素材の種類

| 種類 | 例 |
|---|---|
| テキスト | Instagramキャプションや説明文をそのまま貼り付け |
| 公式サイトURL | `https://gyozamania-honten.com/` |
| 外部メディアURL | 食べログ・Googleマップ等のURL |
| キーワードリスト | ・よだれ鶏<br>・自家製チャーシュー |

### 処理の流れ

```
素材受け取り
  ↓
URLの場合 → WebFetchで内容取得
  ↓
品質チェック（脇役メニュー・情報不足・時制NG の判定）
  ↓
Excelと照合して被りチェック（被りあり / 類似あり / 被りなし）
  ↓
被りなし → お知らせ記事をHTML形式で生成
  ↓
デスクトップの「YYYY年MM月_マニアお知らせ.txt」に追記保存
```

### 出力形式

```
タイトル：〇〇〇〇〇

本文（HTML）：
いつも「[shop-name]」をご利用いただき、誠にありがとうございます。

今回は当店自慢の<strong><span style="color:#c7403a">「メニュー名」</span></strong>をご紹介いたします。
...
<a href="https://[店舗URL]/food" style="color:#c7403a">≫ お料理メニューはこちら</a>
```

---

## 対象店舗

| 店舗名 | サイトURL | ブランドカラー |
|---|---|---|
| 小籠包マニア 中目黒本店 | https://xiaolongbaomania-nakameguro.com/ | `#0f7543` |
| 焼き小籠包マニア | https://yaki-xiaolongbaomania.com/ | `#d61d2a` |
| 北京ダックマニア 虎ノ門ヒルズ本店 | https://pekingduckmania-toranomon.com/ | `#d61d2a` |
| 餃子マニア 虎ノ門ヒルズ店 | https://gyozamania-toranomon.com/ | `#c7403a` |
| 餃子マニア 品川別館 | https://gyozamania-bekkan.com/ | `#c7403a` |
| 餃子マニア 品川本店 | https://gyozamania-honten.com/ | `#c7403a` |

---

## フォルダ構成

```
webmark_mania_news/
├── README.md
├── CLAUDE.md                          # Claude Code 設定
├── news_list/
│   └── get_news_list.py               # 各店舗サイトからお知らせを収集するスクリプト
└── .claude/
    ├── rules/
    │   └── news-checker.md            # 照合ルール
    └── skills/
        └── news-checker/
            └── SKILL.md               # 被りチェック＆記事生成スキル
```

---

## 月次作業手順

### 1. お知らせリストの更新（月1回）

```bash
python news_list/get_news_list.py
```

- 6店舗のサイトから直近2年分のお知らせを自動収集
- 社内NASに `YYYY年MM月_お知らせ内容リスト(2年間).xlsx` として保存
- タスクスケジューラで自動実行設定済み（毎月1日）

### 2. お知らせ記事の作成（都度）

Claude Code でこのフォルダを開き、素材を貼り付けるだけで自動生成されます。

---

## 環境・前提条件

- Python 3.x（`openpyxl` インストール済み）
- 社内NASへのアクセス（`\\192.168.2.241\share\disk1\...`）
- Claude Code（スキル機能使用）
