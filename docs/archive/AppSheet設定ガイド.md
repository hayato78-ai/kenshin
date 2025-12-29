# AppSheet 健診結果入力システム 設定ガイド

## 概要

このガイドでは、健診結果入力システム（労災二次検診・人間ドック・定期検診）のAppSheet設定手順を説明します。

## 前提条件

- Google Workspace アカウント
- AppSheet へのアクセス権限
- 対象スプレッドシートへの編集権限

---

## Step 1: GASセットアップ

### 1.1 スクリプトファイルの追加

GASエディタ（拡張機能 → Apps Script）で以下のファイルを追加:

1. `appSheetSetup.gs` - テーブル作成スクリプト
2. `appSheetWebhook.gs` - Webhookエンドポイント

### 1.2 テーブル作成

1. GASエディタで `setupAppSheetTables` 関数を実行
2. 確認ダイアログで「はい」をクリック
3. 以下のシートが自動作成される:
   - `AS_案件`
   - `AS_受診者`
   - `AS_血液検査`
   - `AS_超音波`
   - `AS_保健指導`
   - `AS_ワークフロー`
   - `AS_所見テンプレート`
   - `AS_判定基準`
   - `AS_設定`

### 1.3 Webhookデプロイ

1. GASエディタ → デプロイ → 新しいデプロイ
2. 種類: ウェブアプリ
3. 次のユーザーとして実行: 自分
4. アクセスできるユーザー: 全員（匿名含む）※社内利用の場合は「組織内の全員」
5. デプロイURLをコピー → `AS_設定` の `WEBHOOK_URL` に設定

---

## Step 2: AppSheet アプリ作成

### 2.1 新規アプリ作成

1. [AppSheet](https://www.appsheet.com/) にアクセス
2. 「+ Create」→「App」→「Start with existing data」
3. 作成したスプレッドシートを選択
4. `AS_案件` シートを選択して「Connect」

### 2.2 テーブル追加

Data → Tables で以下を追加:
- AS_受診者
- AS_血液検査
- AS_超音波
- AS_保健指導
- AS_ワークフロー
- AS_所見テンプレート
- AS_判定基準

---

## Step 3: テーブル設定

### 3.1 AS_案件 (cases)

| カラム | Type | Key | Label | 備考 |
|--------|------|-----|-------|------|
| case_id | Text | ✅ | ✅ | UNIQUEID() |
| case_name | Text | | | |
| exam_type | Enum | | | DOCK, ROSAI_SECONDARY, REGULAR |
| status | Enum | | | 未着手, 処理中, 完了 |
| current_step | Ref | | | → AS_ワークフロー |

**Initial Value設定:**
```
case_id: UNIQUEID()
created_at: NOW()
status: "未着手"
```

### 3.2 AS_受診者 (patients)

| カラム | Type | Key | 備考 |
|--------|------|-----|------|
| patient_id | Text | ✅ | UNIQUEID() |
| case_id | Ref | | → AS_案件 |
| gender | Enum | | M, F |
| status | Enum | | 未入力, 入力中, 確認待ち, 完了 |

**Ref設定:**
- case_id: ReverseName = "受診者一覧"

### 3.3 AS_血液検査 (blood_tests)

| カラム | Type | 備考 |
|--------|------|------|
| blood_test_id | Text (Key) | UNIQUEID() |
| patient_id | Ref | → AS_受診者 |
| *_value | Decimal | 検査値 |
| *_judgment | Enum | A, B, C, D |

### 3.4 AS_超音波 (ultrasound)

| カラム | Type | 備考 |
|--------|------|------|
| ultrasound_id | Text (Key) | UNIQUEID() |
| patient_id | Ref | → AS_受診者 |
| abd_judgment | Enum | A, B, C, D |
| abd_findings | LongText | 自由記述 |

### 3.5 AS_保健指導 (guidance)

| カラム | Type | 備考 |
|--------|------|------|
| guidance_id | Text (Key) | UNIQUEID() |
| patient_id | Ref | → AS_受診者 |
| ai_generated | LongText | AI生成テキスト |
| final_text | LongText | 編集後テキスト |

---

## Step 4: ビュー設定

### 4.1 ダッシュボードビュー

**View name:** Dashboard
**View type:** Dashboard
**Position:** Menu

構成要素:
1. **案件サマリー** (Inline view)
   - 処理中の案件一覧
   - Filter: [status] = "処理中"

2. **クイックアクション** (Action buttons)
   - 新規案件作成
   - CSV取込

3. **本日の予定** (Calendar/Table)
   - 今日の検診案件

### 4.2 案件一覧ビュー

**View name:** 案件一覧
**View type:** Table
**Data:** AS_案件

**Group by:** exam_type
**Sort by:** created_at DESC

**表示カラム:**
- case_name
- exam_type
- status
- patient_count
- completed_count

### 4.3 受診者詳細ビュー

**View name:** 受診者詳細
**View type:** Detail
**Data:** AS_受診者

**Related views:**
- 血液検査 (AS_血液検査)
- 超音波 (AS_超音波)
- 保健指導 (AS_保健指導)

### 4.4 ステップ誘導ビュー

**View name:** 次のアクション
**View type:** Detail
**Display name:** ワークフローナビゲーション

Show if: TRUE (常に表示)

**表示内容:**
```
現在のステップ: <<[current_step].[step_name]>>

次のアクション:
<<[current_step].[step_description]>>

[アクションボタン: 次へ進む]
```

---

## Step 5: アクション設定

### 5.1 CSV取込アクション

**Action name:** CSV取込
**Do this:** App: go to another view within this app
**Target:** CSVアップロードフォーム

### 5.2 保健指導AI生成アクション

**Action name:** AI保健指導生成
**Do this:** Call a webhook
**Webhook URL:** <<[AS_設定].[WEBHOOK_URL]>>

**HTTP Content:**
```json
{
  "action": "generate_guidance",
  "patient_id": "<<[patient_id]>>",
  "case_id": "<<[case_id]>>"
}
```

### 5.3 Excel出力アクション

**Action name:** Excel出力
**Do this:** Call a webhook
**Webhook URL:** <<[AS_設定].[WEBHOOK_URL]>>

**HTTP Content:**
```json
{
  "action": "export_excel",
  "case_id": "<<[case_id]>>",
  "patient_ids": ["<<[patient_id]>>"]
}
```

### 5.4 次のステップへ進むアクション

**Action name:** 次のステップへ
**Do this:** Data: set the value of some columns

**Set these columns:**
- current_step = LOOKUP(
    [current_step].[step_order] + 1,
    AS_ワークフロー,
    step_order,
    step_id
  )
- updated_at = NOW()

---

## Step 6: 自動化設定 (Automation)

### 6.1 判定A自動入力

**Bot name:** 超音波A判定自動入力
**Event:** Data Change → AS_超音波 → Updates Only
**Condition:** [abd_judgment] = "A" AND ISBLANK([abd_findings])

**Task:**
- Set column values: abd_findings = "異常なし"

### 6.2 ステータス自動更新

**Bot name:** 完了ステータス更新
**Event:** Data Change → AS_受診者
**Condition:**
  [blood_test_status] = "済"
  AND [ultrasound_status] IN ("済", "対象外")
  AND [guidance_status] = "済"

**Task:**
- Set column values: status = "完了"

---

## Step 7: フォーマット設定

### 7.1 条件付き書式

**ステータス色分け:**
- 未着手: グレー (#9E9E9E)
- 処理中: 青 (#2196F3)
- 完了: 緑 (#4CAF50)

**判定色分け:**
- A: 緑 (#4CAF50)
- B: 黄緑 (#8BC34A)
- C: オレンジ (#FF9800)
- D: 赤 (#F44336)

### 7.2 バリデーション

**必須チェック:**
- 受診者: name, gender, birth_date
- 血液検査: 各value入力時にjudgmentも必須
- 超音波: judgment入力時

---

## Step 8: テスト

### 8.1 動作確認チェックリスト

- [ ] 案件作成 → 一覧に表示される
- [ ] 受診者追加 → 案件に紐づく
- [ ] CSV取込 → Webhook呼び出し成功
- [ ] 血液検査入力 → 判定自動計算
- [ ] 超音波A判定 → 「異常なし」自動入力
- [ ] AI保健指導生成 → テキスト生成される
- [ ] Excel出力 → ファイル生成される
- [ ] ステータス遷移 → 自動更新される

### 8.2 テストデータ

`AS_案件` にテストデータを追加:
```
case_id: TEST_001
case_name: テスト案件
exam_type: ROSAI_SECONDARY
status: 未着手
```

---

## トラブルシューティング

### Webhook呼び出しエラー

1. GASデプロイURLが正しいか確認
2. デプロイが最新版か確認（新しいデプロイ作成）
3. GASの実行ログでエラー確認

### データが同期されない

1. AppSheet → Data → Regenerate Structure
2. シートのカラム順序がヘッダーと一致しているか確認
3. キーカラムにユニーク値が設定されているか確認

### 判定計算が動かない

1. `AS_判定基準` に該当項目の基準があるか確認
2. gender, exam_type が正しく設定されているか確認

---

## 参考リンク

- [AppSheet ドキュメント](https://support.google.com/appsheet)
- [Apps Script リファレンス](https://developers.google.com/apps-script)
- [Webhook設定ガイド](https://support.google.com/appsheet/answer/10107805)
