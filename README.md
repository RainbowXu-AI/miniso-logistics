# 名创优品物流追踪数据分析看板

## 看板地址
`https://rainbowxu-ai.github.io/miniso-logistics/`

## 如何更新数据
1. 打开本仓库的 `data/` 目录
2. 点击 "Add file" -> "Upload files"
3. 上传最新的 Tracking Report (.xls)
4. 删除旧的 Excel 文件（如有）
5. 点击 "Commit changes"
6. 等待 1-2 分钟，GitHub Action 会自动处理数据
7. 刷新看板页面即可看到最新数据

## 文件结构
- `index.html` - 看板主页面
- `logistics_data.json` - 处理后的数据（由 Action 自动生成）
- `data/` - 存放原始 Excel 文件
- `scripts/process_data.py` - 数据处理脚本
- `.github/workflows/process-data.yml` - 自动处理工作流
