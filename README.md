# Excel 质检工作台

React + TypeScript + Express 的前后端项目，满足以下能力：

- 导入 Excel（`.xls` / `.xlsx`）
- 后端解析表头和行数据
- 支持选择页面展示列
- 强制展示并编辑 `是否合格`、`质检员业务反馈意见`
- 支持 `level1`、`level2` 筛选
- 识别并展示表格中的图片单元格
- 支持多文件导入与左侧导航切换

## 开发

```bash
pnpm install
pnpm dev
```

- 前端: `http://localhost:5173`
- 后端: `http://localhost:8787`

## 构建与类型检查

```bash
pnpm typecheck
pnpm build
```

## 打包与自动同步

```bash
# 默认：打包并自动同步到 https://github.com/iwoov/mySync.git 的 BenchMark 目录
bash scripts/package-zip.sh my-backup.zip

# 只打包，不同步
bash scripts/package-zip.sh my-backup.zip --no-sync

# 自定义本地仓库路径/目标目录
bash scripts/package-zip.sh my-backup.zip --repo-dir /tmp/mySync --target-dir BenchMark
```
