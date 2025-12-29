# 横断的ナレッジ (Cross-Cutting Concerns)

## 概要
このファイルは、3つ以上の機能/領域に影響し、プロジェクト全体の設計方針に関わるナレッジを管理する。
個別のai-guideからは参照リンクのみ設置し、内容の重複を避ける。

## 更新履歴
| 日付 | セッション | 変更内容 |
|------|------------|----------|
| 2024-12-26 | - | 初版作成（テンプレート） |
| 2024-12-28 | - | DP-001 クロスプラットフォーム対応方針を追加 |

---

## 📌 登録基準
以下の**すべて**を満たす場合のみ、このファイルに記載する：
- [ ] 3つ以上の機能/領域に影響する
- [ ] プロジェクト全体の設計方針に関わる
- [ ] 個別のai-guideに記載すると重複が多くなる

**該当しない場合** → 各機能のai-guideに個別記載

---

## 🏗️ 設計方針

### [方針ID: DP-001] クロスプラットフォーム対応方針
**登録日**: 2024-12-28
**関連機能**: python/main.py, python/unified_transcriber.py, python/drive_watcher.py, excel_output
**参照セッション**: -

**方針**:
開発環境（macOS）と使用環境（Windows）の両方で動作するよう、パス設定をプレースホルダー方式で管理する。

**背景・理由**:
- 開発は macOS で行い、実運用は Windows で行うケースがある
- Google Drive for Desktop のパスは OS によって異なる
  - macOS: `/Users/{user}/Library/CloudStorage/GoogleDrive-{email}/マイドライブ/...`
  - Windows: `G:\マイドライブ\...` または `C:\Users\{user}\Google Drive\マイドライブ\...`
- ハードコードされたパスは他の環境で動作しない

**適用ルール**:
1. **パスのハードコード禁止**: Python コード内に絶対パスを直接記述しない
2. **プレースホルダー使用**: `settings_template.yaml` で `${GOOGLE_DRIVE_BASE}` を使用
3. **セットアップスクリプトで置換**: 各環境で `setup.sh` (macOS) / `setup.ps1` (Windows) を実行
4. **設定ファイル参照**: コードは `settings.yaml` からパスを読み込む

**仕組み**:
```
settings_template.yaml          設定テンプレート（プレースホルダー使用）
    │
    ├─ setup.sh (macOS)    ──→  settings.yaml (macOS用パス)
    │
    └─ setup.ps1 (Windows) ──→  settings.yaml (Windows用パス)
```

**プレースホルダー例**:
```yaml
# settings_template.yaml
HUMAN_DOCK:
  template_path: ${GOOGLE_DRIVE_BASE}/50_健診結果入力/05_development/templates/...
```

**生成結果 (macOS)**:
```yaml
# settings.yaml
HUMAN_DOCK:
  template_path: /Users/hytenhd_mac/Library/CloudStorage/GoogleDrive-.../マイドライブ/50_健診結果入力/...
```

**生成結果 (Windows)**:
```yaml
# settings.yaml
HUMAN_DOCK:
  template_path: G:\マイドライブ\50_健診結果入力\...
```

**コードでの読み込み**:
```python
# unified_transcriber.py
config = self.settings.get('exam_types', {}).get('HUMAN_DOCK', {})
self.template_path = Path(config.get('template_path', ''))
```

**例外**:
- `RosaiTranscriber` は現在ハードコードが残存（スコープ外のため未対応）
- 将来的に同じ方式に統一予定

**対応状況**:
| コンポーネント | 対応状況 |
|----------------|----------|
| `HumanDockTranscriber` | ✅ 対応済み |
| `RosaiTranscriber` | ❌ 未対応（スコープ外） |
| `settings_template.yaml` | ✅ プレースホルダー設定済み |
| `setup.sh` (macOS) | ✅ 作成済み |
| `setup.ps1` (Windows) | ✅ 作成済み |

---

### [方針ID: DP-002] （次の方針タイトル）
**登録日**: YYYY-MM-DD
**関連機能**:
**参照セッション**:

**方針**:

**背景・理由**:

**適用ルール**:

**例外**:

---

## 📐 命名規則

### [規則ID: NR-001] （規則タイトル）
**登録日**: YYYY-MM-DD
**関連機能**:

**規則**:

**例**:
```
（コード例やファイル名例）
```

**NG例**:
```
（避けるべき例）
```

---

## ⚠️ エラーハンドリング方針

### [方針ID: EH-001] （方針タイトル）
**登録日**: YYYY-MM-DD
**関連機能**:

**方針**:

**実装パターン**:
```javascript
// コード例
```

---

## 🔗 各ai-guideからの参照方法

各機能のai-guideでは、以下の形式で参照リンクを設置する：

```markdown
## 関連する横断的ナレッジ
- [DP-001] BMLコード標準化方針 → [cross-cutting-concerns.md](../docs/cross-cutting-concerns.md#方針id-dp-001-bmlコード標準化方針)
```

---

## 📋 登録済みナレッジ一覧

| ID | タイトル | 関連機能数 | 登録日 |
|----|----------|------------|--------|
| DP-001 | クロスプラットフォーム対応方針 | 4 | 2024-12-28 |
| （追加時に更新） | | | |

---

## 🗑️ アーカイブ
役割を終えた、または置き換えられたナレッジはここに移動する。

### [ARCHIVED] （旧ナレッジタイトル）
**アーカイブ日**: YYYY-MM-DD
**理由**: （なぜアーカイブされたか）
**置き換え先**: （新しい方針のIDがあれば記載）
