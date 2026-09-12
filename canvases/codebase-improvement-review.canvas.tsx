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

type Priority = "すぐに" | "高" | "中" | "低";

type Finding = {
  id: number;
  priority: Priority;
  score: number;
  items: string;
  feature: string;
  title: string;
  impact: string;
  riskIfUnfixed: string;
  frequency: string;
  tradeoff: string;
  difficulty: string;
  done?: boolean;
};

const FINDINGS: Finding[] = [
  {
    id: 1,
    priority: "高",
    score: 7,
    items: "正確性",
    feature: "データ集約・前段行マージ（結合代入の前処理）",
    title: "行マージと結合比較の正規化の細かい不一致",
    impact: "大半は正確。特定条件でのみ行の横結合がずれる",
    riskIfUnfixed:
      "同一反復内で '001 と 001（や数値と文字列）が混在すると、前段マージだけ別扱いになることあり",
    frequency: "低〜中（キー表記ゆれが混在したときのみ）",
    tradeoff:
      "マージも join_compare に揃えると該当シナリオの行のまとまり方が変わる。反復またぎは引き続き統合しない",
    difficulty: "低",
    done: true,
  },
  {
    id: 2,
    priority: "すぐに",
    score: 10,
    items: "正確性",
    feature: "データ集約・抽出（Excel読取）",
    title: "抽出失敗の except→None/pass が空セルに見える",
    impact: "原因不明の空欄・結合不一致",
    riskIfUnfixed:
      "失敗が空欄に化け、結合失敗や誤った完成表を気付かず確定",
    frequency: "中（破損・書式異常・巨大シート時）",
    tradeoff:
      "失敗をエラー停止にすると一括が途中終了。継続＋明示エラー印の両立が必要",
    difficulty: "中",
    done: true,
  },
  {
    id: 3,
    priority: "すぐに",
    score: 9,
    items: "正確性",
    feature: "データ集約・シート指定読取",
    title: "シート名不一致時の active フォールバック",
    impact: "別シートを黙って読む",
    riskIfUnfixed:
      "指定外シートのデータを集約し、正しそうに見える誤結果を出力",
    frequency: "中（シート名変更・typo時は確実）",
    tradeoff:
      "エラー化すると旧シナリオ（シート名曖昧運用）が止まる。警告付きフォールバック移行も可",
    difficulty: "低",
    done: true,
  },
  {
    id: 4,
    priority: "すぐに",
    score: 9,
    items: "ファイル容量",
    feature: "開発環境（Nuitkaビルドログ）",
    title: "logs/nuitka 約1.5GB の蓄積",
    impact: "約1.5GB即時回収",
    riskIfUnfixed: "ディスク逼迫・バックアップ遅延・ビルド失敗",
    frequency: "高（ビルドするたび増加）",
    tradeoff: "削除すると過去ビルドの原因調査が難しくなる。直近N世代保持が現実的",
    difficulty: "低",
    done: true,
  },
  {
    id: 5,
    priority: "すぐに",
    score: 9,
    items: "ファイル容量",
    feature: "開発環境（配布ビルド成果物）",
    title: "dist/ 旧リリース＋重複ビルド約3.1GB",
    impact: "2GB超削減見込み",
    riskIfUnfixed: "ディスク逼迫・誤配布・旧版混同",
    frequency: "高（リリース積み上げのたび）",
    tradeoff: "旧版ローカル比較ができなくなる。必要版はアーカイブ置き場へ退避",
    difficulty: "低",
    done: true,
  },
  {
    id: 6,
    priority: "すぐに",
    score: 8,
    items: "ファイル容量 / スリム化",
    feature: "リポジトリ衛生（一時ファイル・重複テスト）",
    title: "ルート未追跡ゴミ (_tmp_* / .coverage / 重複テスト)",
    impact: "誤コミット防止＋pytestノイズ排除",
    riskIfUnfixed: "ゴミの誤コミット・テスト収集失敗・レビューノイズ",
    frequency: "中（作業後に残存しやすい）",
    tradeoff: "調査用スクリプトも消える。再現手順が必要なら docs/tools へ正規化",
    difficulty: "低",
    done: true,
  },
  {
    id: 7,
    priority: "高",
    score: 8,
    items: "正確性",
    feature: "データ集約・結合前加工（チェックラベル）",
    title: "加工チェックの部分文字列マッチ",
    impact: "ラベル変更で誤適用",
    riskIfUnfixed:
      "意図しないトリム／日付変換が走り、キー不一致や値破壊",
    frequency: "中（ラベル追加・文言変更時）",
    tradeoff:
      "完全一致＋旧別名許可リスト。部分一致は廃止。安定ID化は将来任意",
    difficulty: "低",
    done: true,
  },
  {
    id: 8,
    priority: "高",
    score: 8,
    items: "正確性 / スリム化",
    feature: "基盤層（core / svc 境界）",
    title: "core → svc 循環依存（A/B 済・host系Cは後回し）",
    impact: "初期化順・Nuitka・単体テストの脆さ",
    riskIfUnfixed: "起動失敗・バンドル抜け・テストが壊れて回帰見逃し",
    frequency: "低〜中（リファクタ・配布ビルド時）",
    tradeoff: "層分割は広い import 差し替えが必要。短期は回帰リスクが上がる",
    difficulty: "中",
    done: true,
  },
  {
    id: 9,
    priority: "高",
    score: 8,
    items: "速度 / スリム化",
    feature: "データ集約コア／デバッグUI全体",
    title: "巨大モノリス分割 (join merge 1継ぎ目済・継続可)",
    impact: "変更リスク局所化",
    riskIfUnfixed: "修正の副作用が広がり、回帰バグ・レビュー不能が常態化",
    frequency: "高（機能追加のたび）",
    tradeoff:
      "分割作業自体が大規模・長期。途中はマージ衝突と二重メンテが増える",
    difficulty: "高",
    done: true,
  },
  {
    id: 10,
    priority: "高",
    score: 7,
    items: "速度",
    feature: "データ集約・Excel抽出（大シート）",
    title: "シート全行 materialize のメモリ/I/O",
    impact: "大シートで数百MB〜数GB・数秒〜数十秒",
    riskIfUnfixed: "メモリ逼迫・ハング・一括処理のタイムアウト",
    frequency: "高（大容量Excel利用時は毎回）",
    tradeoff:
      "範囲限定・ストリーム化は実装複雑。小シートでは全読込の方が単純で速い場合あり",
    difficulty: "高",
    done: true,
  },
  {
    id: 11,
    priority: "高",
    score: 7,
    items: "速度",
    feature: "データ集約・非表示行除外",
    title: "skip_hidden_rows 時の非 read_only 再オープン",
    impact: "読込約2倍",
    riskIfUnfixed: "一括処理が大幅遅延・体感タイムアウト",
    frequency: "高（非表示除外ON時は毎回）",
    tradeoff:
      "read_only維持＋寸法用fullは別開が必要。全シート分を1回のfullでキャッシュし二重化を防ぐ",
    difficulty: "中",
    done: true,
  },
  {
    id: 12,
    priority: "高",
    score: 7,
    items: "正確性 / スリム化",
    feature: "データ集約／UI／更新（横断）",
    title: "広域 except Exception の縮減（シート/非表示ホットパス段階済）",
    impact: "障害可視化・再発防止",
    riskIfUnfixed: "本番障害の原因特定不能・誤った空結果の継続",
    frequency: "高（例外経路は日常的）",
    tradeoff:
      "例外を厳格化するとCOM/Excelの一時失敗で処理が止まりやすくなる",
    difficulty: "中",
    done: true,
  },
  {
    id: 13,
    priority: "高",
    score: 7,
    items: "スリム化 / ファイル容量",
    feature: "旧実装アーカイブ (svc/Old)",
    title: "svc/Old/ 死コード (git追跡)",
    impact: "認知負荷削減",
    riskIfUnfixed: "誤って旧APIを参照・検索ノイズ・レビュー誤誘導",
    frequency: "低（触ったとき）",
    tradeoff: "git履歴・タグに残るが、ローカル参照が不便になる",
    difficulty: "低",
    done: true,
  },
  {
    id: 14,
    priority: "高",
    score: 7,
    items: "スリム化 / 速度",
    feature: "処理サーバ起動 (svc_server)",
    title: "svc_server の DEBUG Popen モンキーパッチ常設",
    impact: "本番ログノイズ除去",
    riskIfUnfixed: "ログ肥大・診断ノイズ・微小オーバーヘッド常時発生",
    frequency: "高（子プロセス起動のたび）",
    tradeoff: "常時OFFだとスポーン調査が手間。環境変数でONにする形が無難",
    difficulty: "低",
    done: true,
  },
  {
    id: 15,
    priority: "中",
    score: 5,
    items: "スリム化",
    feature: "データ集約・polars任意読込",
    title: "_get_polars 三重定義",
    impact: "保守ずれ防止",
    riskIfUnfixed: "片方だけ修正され挙動分岐・再現困難なバグ",
    frequency: "低（修正時のみ）",
    tradeoff: "葉モジュールへ集約。遅延import・失敗時Noneは維持。モジュール統合はしない",
    difficulty: "低",
    done: true,
  },
  {
    id: 16,
    priority: "中",
    score: 5,
    items: "スリム化",
    feature: "日付変換 (YMD / HM)",
    title: "svc_dt_ymd / svc_dt_hm のほぼ複製",
    impact: "数百行削減",
    riskIfUnfixed: "片方だけ直して日付変換結果が機能間で不一致",
    frequency: "低〜中（日付機能改修時）",
    tradeoff:
      "純関数のみ共通化。公開API・config・UI・TOPMOST差は維持（モジュール統合なし）",
    difficulty: "中",
    done: true,
  },
  {
    id: 17,
    priority: "中",
    score: 5,
    items: "速度 / スリム化",
    feature: "データ集約・シナリオ編集UI",
    title: "UIシナリオ編集の過剰 deepcopy",
    impact: "項目多いとUI鈍化",
    riskIfUnfixed: "編集操作のカクつき・入力遅延・誤操作誘発",
    frequency: "中（大規模シナリオ編集時）",
    tradeoff:
      "浅いコピーは副作用で編集内容が汚染される危険。参照管理の設計が要る",
    difficulty: "中",
    done: true,
  },
  {
    id: 18,
    priority: "中",
    score: 5,
    items: "ファイル容量 / スリム化",
    feature: "配布パッケージ／依存管理",
    title: "依存の肥大 (lockと配布の乖離)",
    impact: "数十〜数百MB級の可能性",
    riskIfUnfixed: "配布ZIP肥大・導入時間増・不要DLL混入",
    frequency: "中（配布ビルドごと）",
    tradeoff:
      "polars等を外すと大ファイル高速経路が弱まる。任意依存のままが妥当な場合あり",
    difficulty: "中",
  },
  {
    id: 19,
    priority: "中",
    score: 4,
    items: "ファイル容量",
    feature: "ドキュメント管理",
    title: "ドキュメント版二重・docx/PDF重複",
    impact: "数MB〜10MB",
    riskIfUnfixed: "版の取り違え・説明齟齬・容量浪費",
    frequency: "低（ドキュメント参照時）",
    tradeoff: "docx削除は社内提出物ワークフローと衝突し得る。生成物置き場分離が安全",
    difficulty: "低",
  },
  {
    id: 20,
    priority: "中",
    score: 4,
    items: "速度 / スリム化",
    feature: "データ集約・結合書込ホットパス",
    title: "結合ダンプ／診断コードの本番常駐",
    impact: "ホットパス可読性向上",
    riskIfUnfixed: "保守困難・微小性能ロス・診断フラグ誤ON時の大量出力",
    frequency: "低（通常OFF）／高（誤ON時）",
    tradeoff: "外出しすると現場診断の即応性が下がる。フラグ付き別モジュールが妥協点",
    difficulty: "中",
  },
  {
    id: 21,
    priority: "中",
    score: 4,
    items: "速度",
    feature: "データ集約UI・前面化／ロック再試行",
    title: "ui_data_agg の processEvents / 多重 QTimer",
    impact: "再入・ちらつき低減",
    riskIfUnfixed: "ダイアログちらつき・再入バグ・稀なフリーズ感",
    frequency: "中（前面化・Excel競合時）",
    tradeoff:
      "状態マシン化は実装・検証コスト高。現状の簡易再試行は実装が短い",
    difficulty: "中",
  },
  {
    id: 22,
    priority: "低",
    score: 3,
    items: "スリム化 / ファイル容量",
    feature: "自動テスト／CI",
    title: "テスト肥大・重複シナリオの整理",
    impact: "CI時間短縮",
    riskIfUnfixed: "CI遅延・重複テストのメンテ漏れで偽緑",
    frequency: "中（CI実行のたび）",
    tradeoff: "統合しすぎるとカバレッジ穴。回帰の重要ケースは残す必要がある",
    difficulty: "中",
  },
];

function priorityLabel(f: Finding): string {
  if (f.priority === "すぐに") return `すぐに ${f.score}`;
  return `${f.priority}${f.score}`;
}

function cellText(f: Finding, value: string) {
  if (!f.done) return value;
  return (
    <Text size="small" style={{ textDecoration: "line-through" }} tone="secondary">
      {value}
    </Text>
  );
}

export default function CodebaseImprovementReview() {
  // スコア順を維持（完了行を末尾へ送らない。見え消しは現状位置の文言に付与）
  const sorted = FINDINGS.slice().sort((a, b) => b.score - a.score || a.id - b.id);
  const doneCount = FINDINGS.filter((f) => f.done).length;

  return (
    <Stack gap={16}>
      <H1>コードモジュール見直し — 改善一覧</H1>
      <Text tone="secondary" size="small">
        Excel_AddIn 全体レビュー / A・B・C・D 修正済 / 見え消し＝対応済（行位置はスコア順のまま）
      </Text>

      <Grid columns={4} gap={12}>
        <Stat value={String(doneCount)} label="修正済み" tone="success" />
        <Stat value="0" label="すぐ残" tone="danger" />
        <Stat value="9" label="高優先度" tone="warning" />
        <Stat value="7" label="正確性関連" tone="info" />
      </Grid>

      <Callout tone="success" title="A / B / C / D / #7 / #15+#16 修正完了">
        正確性・掃除・I/O・加工チェックに加え、#15 polars 遅延読込の単一化、#16
        日付変換の共通純関数抽出（YMD/HM の公開API・config差は維持）まで完了。
      </Callout>

      <Callout tone="warning" title="#11 速度回帰の追加修正（2026-09-11）">
        D 初回実装で「シートごとに full 再オープン」となり実機が約2倍遅延。read_only
        抽出キャッシュは維持したまま、寸法用 full をファイルあたり1回にし全シートの非表示行を
        一括キャッシュするよう是正（旧速度水準への復帰目的）。
      </Callout>

      <Callout tone="info" title="#7 合意仕様（実装済）">
        現行3ラベルは完全一致。旧シナリオ文言（日付変換…／全角→半角（英数字・記号）／半角変換
        等）も同じ加工として許可。部分文字列マッチは廃止。UI復元と実行時で同一別名集合を共有。
      </Callout>

      <Callout tone="info" title="#1 合意仕様（実装済）">
        比較は文字列＋先頭 ' 除去。反復またぎは統合しない。結合代入本線は従来どおり。
      </Callout>

      <Callout tone="success" title="構造改善フェーズ完了（2026-09-12）">
        #8A/#8B / #17 / #12（ホットパス段階）/ #9（join merge 1継ぎ目）を実施。復帰点:
        backup/pre-structure-improve-20260912（be555b5）。詳細は
        structure-improve-phase-report Canvas。
      </Callout>

      <Callout tone="warning" title="最優先の結論">
        正確性止血と掃除・I/O・今回の構造改善（段階）は完了。残は #8C（host循環）、
        #9 の追加継ぎ目、#12 の他ホットパス、#18〜 など。
      </Callout>

      <H2>改善優先度一覧（スコア降順）</H2>
      <Table
        striped
        stickyHeader
        headers={[
          "#",
          "優先度",
          "改善アイテム",
          "対象機能名",
          "改善内容",
          "改善度合い",
          "未改善時に起こり得る事象",
          "頻度",
          "トレードオフ",
          "難易度",
        ]}
        columnAlign={[
          "right",
          "left",
          "left",
          "left",
          "left",
          "left",
          "left",
          "left",
          "left",
          "center",
        ]}
        rowTone={sorted.map((f) =>
          f.priority === "すぐに"
            ? "danger"
            : f.priority === "高"
              ? "warning"
              : f.priority === "中"
                ? "info"
                : "neutral",
        )}
            rows={sorted.map((f) => [
              cellText(f, String(f.id)),
              cellText(f, (f.done ? "済 " : "") + priorityLabel(f)),
              cellText(f, f.items),
              cellText(f, f.feature),
              cellText(f, f.title),
              cellText(f, f.impact),
              cellText(f, f.riskIfUnfixed),
              cellText(f, f.frequency),
              cellText(f, f.tradeoff),
              cellText(f, f.difficulty),
            ])}
      />

      <Divider />

      <H2>推奨着手順・抱き合わせ</H2>
      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>完了済み</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>A(#2+#3) / B / C(#1) / D(#10+#11) / #7 / #15+#16 → 済</Text>
              <Text>#8A+#8B / #17 / #12段階 / #9 join-merge1 → 済（2026-09-12）</Text>
              <Text>#10 追加ストリーム化は見送り（多ファイル小容量向けには効果薄）</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>次の着手（未着手）</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>次候補: #18〜 / #8C（host） / #9 追加継ぎ目 / #12 他ホットパス</Text>
              <Text>#8C・#9 残りは単独・長期（キャッシュ寿命を壊さないこと）</Text>
              <Text>復帰: git reset --hard backup/pre-structure-improve-20260912</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H2>#1 合意仕様（要約）</H2>
      <Table
        headers={["項目", "内容"]}
        rows={[
          ["連携キー", "同一パス＋ファイル＋シート内で値を横連携"],
          [
            "結合キー",
            "ファイル／シートをまたぎ、指定項目と座標値を比較し一致行へ主キー代入。複数一致は全部。主キー順で後勝ち",
          ],
          ["比較", "値は文字列扱い。先頭 ' を除き '001 と 001 は同一"],
          [
            "#1の修正範囲",
            "実装済: 前段 _merge_rows_by_join_keys を join_compare に揃え（反復またぎは統合しない）",
          ],
        ]}
      />

      <H2>頻度の読み方</H2>
      <Table
        headers={["表記", "意味"]}
        rows={[
          ["高", "該当機能利用時にほぼ毎回、または日常運用で頻発"],
          ["中", "特定条件（大ファイル・設定ON・改修時など）で発生"],
          ["低", "稀、または触ったとき／誤設定時のみ"],
        ]}
      />

      <Text size="small" tone="secondary">
        頻度・トレードオフは実測テレメトリではなく、コード経路と利用想定からの試算です。Canvas
        を開き直すと最新版が表示されます。
      </Text>
    </Stack>
  );
}
