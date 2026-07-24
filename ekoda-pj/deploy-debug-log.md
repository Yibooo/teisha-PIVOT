# デプロイ不具合デバッグ記録（teisha-navi / Vercel + Convex）

駅徒歩フィルター実装中に発覚した「Vercel本番デプロイが失敗し続ける」問題の調査記録。**機能コードの不具合ではなく、デプロイ基盤（ビルド設定）の問題**だった。

## 症状
- Vercel で teisha-navi が **「No Production Deployment」**。直近の複数コミットが連続でビルド失敗。
- 本番サイト `teisha-navi.vercel.app` は**古い成功ビルドのまま**動作（最近の変更が未反映）。
- ビルドエラー：
  ```
  Module not found: Can't resolve '@/convex/_generated/api'  （全ページ・11エラー）
  Error: Command "npm run build" exited with 1
  ```

## 根本原因
1. `convex/_generated/` は **.gitignore**（＝Gitに含めず、ビルド時に生成が必要）。
2. Vercel の **Build Command が `next build` のみ**で、Convexの生成（codegen/deploy）を実行していなかった。
   → `_generated` が存在しないまま `next build` が走り、全ページのimportが解決できず失敗。
3. さらに **`CONVEX_DEPLOY_KEY` が Vercel 環境変数に無かった**ため、Build Command を `npx convex deploy --cmd 'npm run build'` に変えても、convex deploy が「鍵なし」と判断してデプロイ工程をスキップ → 同じ失敗。

### 切り分けの決め手
- ローカルでは `npx convex dev --once --local` で `_generated` を生成すれば `next build` は**成功**（tsc エラー0）。→ アプリコードは正常。
- 失敗ログで `convex deploy` が「Deploying to ...」を出さず、1.3秒でいきなり `npm run build` に進んでいた。→ 鍵未検出でデプロイをスキップしていた証拠。

## 対応（恒久修正）
1. Vercel → Settings → Build and Deployment → **Build Command** を Override ON で
   ```
   npx convex deploy --cmd 'npm run build'
   ```
2. Convex（本番デプロイ **precious-mongoose-658**）の Settings → Deploy Keys → **Production Deploy Key を発行**
   - ※本番Convexは `precious-mongoose-658.convex.cloud`（初期に誤って `happy-otter-123` と推定したが、それは古いバンドルの値だった）
3. Vercel → **Environment Variables** に `CONVEX_DEPLOY_KEY = prod:...` を追加（Production + Preview）
4. Deployments → Redeploy（キャッシュ無し）

## 学び / 注意点
- Convex + Vercel は「Build Command で `convex deploy --cmd` を実行」する構成が前提。`next build` 単体では `_generated` が無くて必ず失敗する。
- デプロイキーは **Vercel** の環境変数に入れる（Convex側の env に入れても無意味）。

## 方針転換（重要）
駅徒歩フィルターの実装アプローチを、**バックエンド再集計 → フロントエンドのみのフィルターに変更**。
- 理由：エリア（区）フィルターと同様、各物件が既に持つ「駅徒歩分」で**一覧を絞るだけ**で目的を達成できる。バックエンドの集計改修は過剰だった。
- 新方針：`getPropertiesWithDataCounts` が返す `walkMinutes` を使い、**物件一覧テーブルをクライアント側でフィルタ**。バックエンド関数は変更しない（＝デプロイリスク最小）。
- 上部の2グラフはエリア単位のまま（駅徒歩フィルターは物件一覧に適用、と注記）。

_記録日：2026-07-24_
