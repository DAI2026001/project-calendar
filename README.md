# 投标日程 H5

这是一个可直接部署到 Vercel 的静态 H5 页面。

## 文件说明

- `index.html`：日历页面
- `schedule-data.js`：从 Excel 生成的项目数据
- `build_schedule.py`：把 Excel 转成前端数据的脚本
- `vercel.json`：Vercel 静态部署配置

## 本地更新数据

```bash
python3 build_schedule.py /Users/sdai/Downloads/投标日程表.xlsx ./schedule-data.js
```

## 部署到 Vercel

1. 将当前目录上传到 GitHub 仓库
2. 在 Vercel 中导入该仓库
3. Framework Preset 选择 `Other`
4. Build Command 留空
5. Output Directory 留空
6. 点击 Deploy

部署完成后会自动生成一个 HTTPS 链接，可直接在微信中转发打开。
