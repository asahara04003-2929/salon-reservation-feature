# GAS → ローカルに反映（最新を取る）
cd gas
clasp pull

# ローカル → GASに反映（デプロイは行わない。単純のソースコードの反映のみ）
cd gas
clasp push

# 新しいデプロイを作る（=新しいversionを発行してデプロイ）
clasp deploy -m "update"

# よくあるハマりポイント（先回り）
- A) clasp clone で別フォルダが増える/思った場所に来ない
- --rootDir . を付けると “今のフォルダ直下” に来ます。

- B) appsscript.json がない/見えない
- clasp clone すると入るはず。見えない場合は隠し設定ではなく、clone場所違いが多いです。

- C) TypeScript/モダン構成にしたい
- 今は「まず同期してGit管理」が目的でOK。必要なら clasp + webpack / esbuild でビルドしてpushする構成に拡張できます。