import {
  Callout,
  Card,
  CardBody,
  CardHeader,
  Divider,
  Grid,
  H1,
  H2,
  Stack,
  Stat,
  Table,
  Text,
} from "cursor/canvas";

export default function StructureImprovePhaseReport() {
  return (
    <Stack gap={20}>
      <H1>構造改善フェーズ報告（#8A→#17→#12→#8B→#9）</H1>
      <Text tone="secondary">
        実施日 2026-09-12。挙動変更を最小化し、改善→テストをステップ実行。速度に影響しうる
        キャッシュ寿命は未変更。
      </Text>

      <Grid columns={4} gap={12}>
        <Stat value="5" label="完了フェーズ" tone="success" />
        <Stat value="123" label="通し関連 PASS" tone="success" />
        <Stat value="1" label="既知 FAIL（既存）" tone="warning" />
        <Stat value="be555b5" label="復帰コミット" />
      </Grid>

      <Callout tone="info" title="いつでも戻す">
        ブランチ backup/pre-structure-improve-20260912 = 構造改善前チェックポイント
        （be555b5）。戻し方: git reset --hard backup/pre-structure-improve-20260912
      </Callout>

      <H2>各フェーズ結果</H2>
      <Table
        striped
        headers={["Phase", "内容", "変更の要点", "テスト", "速度・波及"]}
        columnAlign={["left", "left", "left", "left", "left"]}
        rows={[
          [
            "#8A",
            "循環依存切断",
            "coerce を core_excel_text へ。core_join_compare は svc 非依存",
            "19 passed（join/link/merge/value_post）",
            "ホットパス同等（関数移動のみ）",
          ],
          [
            "#17",
            "UI deepcopy",
            "複製名ループの deepcopy 廃止。登録 validate の全 items deepcopy 縮減。get_item の deepcopy は維持",
            "3 passed（scenario deepcopy）+ layout helpers",
            "大規模シナリオの複製・登録が軽くなる想定。抽出速度は無関係",
          ],
          [
            "#12",
            "except 縮減（段階）",
            "シート名一覧・非表示行 I/O の Exception を想定型へ。失敗時 warning ログ",
            "17 passed（except/io/extract/hidden）",
            "想定 I/O 失敗は従来どおり []/set()。想定外は伝播（黙殺減）",
          ],
          [
            "#8B",
            "path_network",
            "実装を core_path_network へ。svc は re-export。core_env は core 参照",
            "22 passed（network/stage/io_profile）",
            "判定ロジック同一",
          ],
          [
            "#9",
            "モノリス分割 1継ぎ目",
            "data_agg_join_merge へ merge/norm 抽出。svc_data_agg は re-import。cache 非接触",
            "32 passed（merge + join_search）",
            "キャッシュ寿命不変。結合探索本体は未分割",
          ],
        ]}
      />

      <H2>全体通し確認</H2>
      <Card>
        <CardHeader>関連スイート（venv python）</CardHeader>
        <CardBody>
          <Stack gap={6}>
            <Text>
              結果: 123 passed / 1 failed（test_compute_batch_parallel_matches_sequential）
            </Text>
            <Text>
              失敗は構造改善前チェックポイント（be555b5）でも再現。precache が別ファイル専用シート
              「紐付け履歴」を現ブックで解決しようとして DataAggSheetMissingError（#2/#3 系の既存課題）。
              本フェーズ起因のデグレではない。
            </Text>
            <Text>
              構造整合: core_join_compare / core_env の svc 即時辺なし。merge は
              data_agg_join_merge と svc_data_agg で同一参照。
            </Text>
          </Stack>
        </CardBody>
      </Card>

      <Callout tone="warning" title="既知・範囲外（実機前に把握）">
        test_data_agg_check_labels の日付ラベル2件は shape_date_value がシリアルを日付化しない既存不具合（本フェーズ非改変）。#8C host 循環・#9 追加分割・#12 他ホットパスは未実施。
      </Callout>

      <Divider />

      <H2>実機確認の観点（ご担当）</H2>
      <Table
        headers={["観点", "確認内容"]}
        rows={[
          ["シナリオ編集", "ソース複製・登録・Undo。大規模シナリオでカクつき改善感"],
          ["一括／ステップ", "結合・連携・シート条件。速度が明らかに悪化していないこと"],
          ["非表示行除外", "skip_hidden_rows ON で結果と所要時間が従来水準"],
          ["ネットワークパス", "UNC 起点のステージング／並列ランプが従来どおり"],
          ["不都合時", "backup/pre-structure-improve-20260912 へ reset して切り戻し判断"],
        ]}
      />

      <Text size="small" tone="secondary">
        Source: ローカル pytest（.venv）・git backup/pre-structure-improve-20260912 · 2026-09-12
      </Text>
    </Stack>
  );
}
