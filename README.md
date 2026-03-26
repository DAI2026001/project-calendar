# 投标日程 H5

这是一个可直接部署到 Vercel 的静态 H5 页面。

## 文件说明

- `index.html`：日历页面
- `schedule-data.js`：从 Excel 生成的项目数据
- `build_schedule.py`：把 Excel 转成前端数据的脚本
- `vercel.json`：Vercel 静态部署配置

## 本地更新数据

```bash
python3 build_schedule.py /Users/sdai/Downloads/辉阳集团项目.xlsx ./schedule-data.js
```

也可以直接使用固定文件名的一键更新脚本：

```bash
zsh /Users/sdai/document/代码/condex/update_schedule.sh
```

使用方式：

1. 在腾讯文档中导出最新 Excel
2. 保存并覆盖 `/Users/sdai/Downloads/辉阳集团项目.xlsx`
3. 执行上面的脚本
4. 如需同步到线上，再执行 `git add / git commit / git push`

## 部署到 Vercel

1. 将当前目录上传到 GitHub 仓库
2. 在 Vercel 中导入该仓库
3. Framework Preset 选择 `Other`
4. Build Command 留空
5. Output Directory 留空
6. 点击 Deploy

部署完成后会自动生成一个 HTTPS 链接，可直接在微信中转发打开。
