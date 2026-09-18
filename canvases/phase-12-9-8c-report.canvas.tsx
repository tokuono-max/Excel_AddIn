import {
  Callout,
  Divider,
  H1,
  H2,
  Stack,
  Table,
  Text,
} from "cursor/canvas";

export default function Phase1298cReport() {
  return (
    <Stack gap={16}>
      <H1>フェーズ報告 — #12 → #9 → #8C</H1>
      <Text tone="secondary" size="small">
        2026-09-12。この段は結合有無・連携先集約・ファイル通過判定の移動。xlsx キャッシュ・抽出の照合・COM は未変更。Nuitka は未実施。実機は未確認。
      </Text>

      <H2>各フェーズ</H2>
      <Table
        striped
        headers={["フェーズ", "変更", "テスト"]}
        rows={[
          [
            "#12",
            "読めないセルの値を #ERR_EXTRACT に変更。指定シート無しは空スキップのまま。xls 読取 import は ImportError / AttributeError のみ",
            "test_data_agg_extract_sheet_and_error_mark.py PASS",
          ],
          [
            "#9",
            "結合の有無、連携先の集約、ファイル通過判定を source_ui へ移設。判定は同じ。再エクスポートで呼び出し側は維持",
            "test_data_agg_source_pattern_seam.py / test_data_agg_join_search.py / test_data_agg_link_write_and_match_keys.py PASS",
          ],
          [
            "#8C",
            "この段では起動を動かしていない",
            "対象外",
          ],
        ]}
      />

      <H2>通し確認</H2>
      <Table
        striped
        headers={["対象", "結果"]}
        rows={[
          ["この段の関連 + public API + phase dividers + 結合検索 + 連携書き込み", "45 passed（進捗文言3件を除く）"],
          [
            "test_data_agg_debug_step_snapshots.py の進捗文言 3件",
            "FAIL。集約結果ではない。本キャンペーン終了後にテストを今の画面へ合わせる",
          ],
          [
            "実機（この段）",
            "未確認。ビルド不要。いつもの2シナリオが前回と同じ行数で完走すればよい",
          ],
          ["Nuitka / 差分インストール", "未実施（全改善の最後まで保留）"],
        ]}
      />

      <Callout tone="warning" title="わざと残したもの">
        COM の広域 except と Nuitka は動かしていない。xlsx キャッシュと抽出の照合も動かしていない。
      </Callout>

      <Divider />
      <Text size="small" tone="secondary">
        この段はテストまで。実機確認待ち。ビルドは不要。進捗文言テスト3件は既知で、この段の対象外。
      </Text>
    </Stack>
  );
}
