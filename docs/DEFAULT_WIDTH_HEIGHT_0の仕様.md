# DEFAULT_WIDTH / DEFAULT_HEIGHT が 0 の場合の仕様

**目的**: 全画面で `DEFAULT_WIDTH` および `DEFAULT_HEIGHT` を 0（または未指定）にした場合、**オートサイズ（コンテンツに合わせた枠サイズ）** になることを保証する。

---

## 仕様（0 = オートサイズ）

- **0 または未指定**: ウィンドウは `adjustSize()` および `sizeHint()` に基づき、コンテンツに合わせたサイズに自動調整される。
- **正の値**: 指定した幅・高さ（ピクセル）でウィンドウをリサイズする。

---

## 画面別の実装状況

| 画面・モジュール | 設定経路 | 0 の扱い | 備考 |
|------------------|----------|----------|------|
| **apply_window_config を利用する画面** | 各 JSON の WINDOW | ✅ 0 = オート | リサイズ不可時: adjustSize + setFixedSize(sizeHint)。リサイズ可時: 0 なら resize しない（レイアウト任せ）。 |
| ワーニング（ui_common） | SCREENS.WARNING.WINDOW | ✅ 修正済み | 従来 0→420x140 のフォールバックだったが、0 のとき adjustSize + resize(sizeHint) に変更。 |
| 進捗（ProgressDialog） | apply_window_config | ✅ 0 = オート | 共通 apply に準拠。 |
| 完了（DoneDialog） | apply_window_config | ✅ 0 = オート | 同上。 |
| フォルダ選択（ui_fld） | WINDOW | ✅ 0 = オート | def_w/def_h が両方 0 のときは resize しない（レイアウトの自然サイズ）。片方のみ 0 のときは sizeHint で補う。 |
| 分割（csv_sp MAIN） | MAIN.WINDOW マージ | ✅ 修正済み | 従来 0→560x420 のフォールバックだったが、0 のとき adjustSize + sizeHint を使用。不足時のみ 560x420 をフォールバック。 |
| 結合（csv_mg） | apply_window_config | ✅ 0 = オート | 共通 apply に準拠。 |
| 行整形（hd_nr） | apply_window_config | ✅ 0 = オート | 共通 apply に準拠。 |
| その他（csv_ld, csv_sv, undo 等） | 各 JSON WINDOW | ✅ 0 = オート | apply_window_config を呼ぶ画面はすべて同じ仕様。 |

---

## 修正内容（2026-03-09）

1. **ui_common.py**  
   ワーニングダイアログで `DEFAULT_WIDTH` / `DEFAULT_HEIGHT` が 0 のとき、`resize(420, 140)` を行わず、`adjustSize()` と `resize(sizeHint())` でオートサイズとするよう変更。

2. **ui_csv_sp.py**  
   分割メイン画面で幅・高さのいずれかが 0 のとき、即 560x420 にせず、`adjustSize()` と `sizeHint()` で幅・高さを決めるよう変更。sizeHint が 0 の場合のみ 560x420 をフォールバック。

3. **apply_window_config**  
   コメントで「0 または未指定 = オートサイズ（adjustSize/sizeHint）」であることを明記。

---

## JSON での指定例

```json
"WINDOW": {
  "DEFAULT_WIDTH": 0,
  "DEFAULT_HEIGHT": 0
}
```

上記のように 0 を指定すると、その画面はコンテンツに合わせた自動枠サイズで表示される。
