# 【依頼書】GitHub App「Claude」へのリポジトリアクセス権限付与について

| 項目 | 内容 |
|---|---|
| 起票日 | 2026-08-26 |
| 依頼者 | scarab@superior.co.jp |
| 対応者 | GitHub Organization `fujiwara-superior` の Owner（または GitHub App 管理権限保有者） |
| 対象リポジトリ | **`fujiwara-superior/-`**（リポジトリ名はハイフン1文字） |
| 緊急度 | 中（作業成果物がリモートに保存できない状態） |
| 想定作業時間 | 5分程度 |

---

## 1. 発生している事象

Claude Code から対象リポジトリへ `git push` を実行すると、**HTTP 403** で失敗します。

**実際のエラー出力（原文）：**

```
remote: Claude doesn't have GitHub access to fujiwara-superior/- for your organization.
        An org admin can install the Claude GitHub App at
        https://github.com/apps/claude/installations/select_target, or reconnect GitHub
        from claude.ai settings to re-link an existing installation
fatal: unable to access 'https://github.com/fujiwara-superior/-/': The requested URL returned error: 403
```

**影響：**
ローカルでのコミットは成功していますが、リモートへ反映できません。

- ブランチ：`claude/showa-spring-proposal-euikzw`
- 未 push のコミット：`183d07c`（営業提案用ドキュメント 1ファイルの追加のみ。アプリケーションコードの変更なし）

---

## 2. 原因

**GitHub App「Claude」が、対象リポジトリへのアクセスを許可されていない**ことが原因です。

次のいずれかの状態にあります。

1. Organization `fujiwara-superior` に GitHub App「Claude」がインストールされていない
2. インストール済みだが、Repository access の対象に `-` リポジトリが**含まれていない**

**認証情報の誤り・コミット署名・ネットワークの問題ではありません。** 権限設定のみで解消します。

> ※ 補足：`git push` にはリポジトリの **Contents（コード）への書き込み権限**が必要です。App のインストール画面に表示される権限をそのまま承認してください。

---

## 3. 対応手順

### 手順A（本命）— GitHub App のインストール／対象リポジトリの追加

1. 次の URL を開く
   **https://github.com/apps/claude/installations/select_target**

2. インストール先として **`fujiwara-superior`**（Organization）を選択する
   - 個人アカウントではなく、**Organization を選ぶ**点にご注意ください。

3. Repository access を設定する
   - **All repositories**（全リポジトリを許可）
     または
   - **Only select repositories** を選び、リポジトリ **`-`** を明示的に追加

   > ⚠️ **要注意**：対象リポジトリ名は**ハイフン1文字（`-`）**です。
   > リポジトリ選択のプルダウンでは表示幅が狭く見落としやすいため、追加後に
   > 一覧へ `-` が表示されていることを必ず目視確認してください。

4. 表示された権限内容を確認し、**Install**（初回）または **Save**（既にインストール済みの場合は Configure 画面）をクリック

---

### 手順B — Claude 側の組織設定でリポジトリが許可対象か確認

管理画面から、対象リポジトリが許可リストに含まれているかを確認してください。

**https://claude.ai/admin-settings/claude-tag**

`fujiwara-superior/-` が許可対象に入っていない場合は、追加してください。

---

### 手順C — 手順A・B を実施しても解消しない場合（利用者側の作業）

既存インストールとの紐付けが切れている可能性があります。利用者（scarab@superior.co.jp）側で GitHub 連携を再実行します。

**https://claude.ai/customize/connectors?auth_start=github&auth_start_force=1**

※ この手順は管理者ではなく**利用者本人**が実施します。手順A・B完了の連絡をいただいた後に、こちらで実施します。

---

## 4. 完了確認の方法

**管理者側での確認：**

- GitHub → Organization `fujiwara-superior` → Settings → GitHub Apps → **Claude** → Configure
- Repository access の欄に **`-`** が表示されている（または「All repositories」が選択されている）

**依頼者側での確認：**

- 新しい Claude Code セッションを開始し、`git push` を再実行 → 403 が発生しないこと

> ※ セッションは一時的な実行環境のため、権限付与後にセッションを開き直す必要がある場合があります。権限反映の確認はこちらで行います。

---

## 5. 完了時にご連絡いただきたい内容

以下を依頼者へお知らせください。

1. 実施した手順（A / B / C のいずれか、複数可）
2. Repository access の設定（**All repositories** か **Only select repositories** か）
3. 実施日時
4. 実施の際にエラー・警告が表示された場合はその内容

---

## 6. 補足事項

- 本件で push する内容は、顧客提案用の Markdown ドキュメント 1ファイルのみです。既存のコード・設定への変更は含まれません。
- Repository access を **Only select repositories** に絞る運用でも問題ありません。その場合、今後 Claude Code で扱うリポジトリが増えるたびに同様の追加作業が必要になります。
- 上記以外の URL・設定画面からの操作は不要です。不明点があれば実施前に依頼者までご確認ください。
