import {
  Callout,
  Divider,
  H1,
  H2,
  Stack,
  Stat,
  Grid,
  Table,
  Text,
  BarChart,
} from "cursor/canvas";

export default function PerfRegression1822Log() {
  return (
    <Stack gap={20}>
      <H1>実機遅延のログ解析（#18/#22 後）</H1>
      <Text tone="secondary">
        Source: hc_csv.log · compute_batch_timing · カットオフ 2026-09-12 11:40 ·
        UI秒はスクショ、内部は total_ms / extract_ms
      </Text>

      <Callout tone="warning" title="結論">
        遅延は事実（両シナリオ）。内訳はほぼ extract（抽出並列）。merge/join/table
        は不変。#18/#22（requirements・テスト配置）はホットパスに無い。11:33
        Excel終了→11:42 svc 再起動の直後計測。ODN-164 は3回目でほぼ復帰、ODN375
        は +約5秒で安定して遅い。原因切り分けは未コミットの抽出系差分の A/B が先。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="+7.3s" label="ODN-164 total平均" tone="warning" />
        <Stat value="+5.3s" label="ODN375 total平均" tone="warning" />
        <Stat value="~0s" label="merge/table差分" tone="success" />
        <Stat value="extract" label="主因フェーズ" tone="danger" />
      </Grid>

      <H2>ログ上の compute total（秒）</H2>
      <BarChart
        categories={["ODN164前", "ODN164後", "ODN375前", "ODN375後"]}
        series={[
          {
            name: "avg total_ms",
            data: [45.13, 52.39, 15.64, 20.95],
            tone: "warning",
          },
        ]}
        height={220}
      />
      <Text size="small" tone="secondary">
        ODN-164: BEFORE n=10 / AFTER n=3 · ODN375: BEFORE n=2 / AFTER n=2
      </Text>

      <H2>内訳（平均）</H2>
      <Table
        striped
        headers={[
          "シナリオ",
          "区分",
          "total平均",
          "extract平均*",
          "merge",
          "table",
          "UIスクショ目安",
        ]}
        columnAlign={[
          "left",
          "left",
          "right",
          "right",
          "right",
          "right",
          "right",
        ]}
        rows={[
          ["ODN-164", "前", "45.1s", "206s", "~0.2s", "~1.1s", "~49s"],
          ["ODN-164", "後", "52.4s", "234s", "~0.2s", "~1.3s", "52–60s"],
          ["ODN375", "前", "15.6s", "101s", "~0.01s", "~0.03s", "~18s"],
          ["ODN375", "後", "21.0s", "133s", "~0.01s", "~0.03s", "23–24s"],
        ]}
      />
      <Text size="small" tone="secondary">
        *extract_ms は並列ワーカー合計に近い累積で、壁時計の total_ms より大きい。
        差分の向き（後の方が大きい）が重要。
      </Text>

      <H2>ODN-164 後の推移（ウォームアップ疑い）</H2>
      <Table
        striped
        headers={["時刻", "total_ms", "extract_ms", "UI秒"]}
        columnAlign={["left", "right", "right", "right"]}
        rows={[
          ["11:44:19", "56.2s", "252s", "59.98"],
          ["11:45:25", "52.8s", "232s", "56.29"],
          ["11:46:58", "48.1s", "218s", "51.57"],
          ["参考: 11:23前", "45.6s", "218s", "49.09"],
        ]}
      />

      <H2>プロセス境界</H2>
      <Table
        striped
        headers={["時刻", "事象"]}
        rows={[
          ["11:22", "svc pid=25556 起動 → 11:23 ODN-164（速い側）"],
          ["11:33", "Excel終了 → svc/ui shutdown"],
          ["11:33–11:42", "#18/#22 作業ウィンドウ（requirements/tests）"],
          ["11:42", "svc pid=1872 再起動 → 以降が遅い側の計測"],
        ]}
      />

      <H2>#18/#22 との因果</H2>
      <Callout tone="info" title="ホットパス未変更">
        requirements 分割・テスト衛星統合のみ。polars 削除なし。ログにも polars
        失敗・join_dump ON・workers 変更なし（常時 workers=8）。
      </Callout>

      <H2>次の切り分け（推奨順）</H2>
      <Table
        striped
        headers={["#", "手順", "期待"]}
        rows={[
          [
            "1",
            "同じ Excel セッションで ODN375 をさらに2回（ウォーム確認）",
            "落ちればキャッシュ要因、残れば回帰",
          ],
          [
            "2",
            "未コミットの svc_data_agg_extract.py（#12）だけ一時退避して再計測",
            "戻れば #12 系が主因",
          ],
          [
            "3",
            "それでも遅ければ svc_data_agg.py / join_dump 分離も同様に A/B",
            "構造改善側の切り分け",
          ],
        ]}
      />

      <Divider />
      <Text size="small" tone="secondary">
        差分インストールは不要。ビルド検証も本遅延の切り分けには不要。
      </Text>
    </Stack>
  );
}
