#!/usr/bin/env bash
set -euo pipefail

# 用途：在网络不稳定时推送 feature 分支；失败时给出 bundle 兜底方案
# 用法：scripts/push_feature_branch.sh feature/process-graph-aggregation

BRANCH="${1:-feature/process-graph-aggregation}"
REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
BUNDLE_PATH="$HOME/${BRANCH//\//-}.bundle"

cd "$REPO_DIR"

echo "[info] 当前仓库: $REPO_DIR"
echo "[info] 待推送分支: $BRANCH"

# 确保本地有该分支
if ! git show-ref --verify --quiet "refs/heads/$BRANCH"; then
  echo "[error] 本地不存在分支 $BRANCH" >&2
  exit 1
fi

set +e
git push -u origin "$BRANCH"
code=$?
set -e

if [ $code -eq 0 ]; then
  echo "[ok] 推送成功: $BRANCH"
  exit 0
fi

echo "[warn] 推送失败，正在创建离线 bundle 以便在其他环境推送..."
git bundle create "$BUNDLE_PATH" "$BRANCH"
echo "[ok] 已创建 bundle: $BUNDLE_PATH"
cat <<EOF

在另一台网络可用的机器上可执行以下步骤：

  # 方式A：从 bundle 克隆
  git clone "$BUNDLE_PATH" repo-from-bundle
  cd repo-from-bundle
  git checkout $BRANCH
  git remote add origin <your-remote-url>
  git push -u origin $BRANCH

  # 方式B：导入到现有仓库
  cd /path/to/your/repo
  git fetch "$BUNDLE_PATH" $BRANCH:$BRANCH
  git checkout $BRANCH
  git push -u origin $BRANCH

EOF

exit 2
