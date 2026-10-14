# 证照印章管理系统（前端）

本仓库仅存放前端与后端源码，**不包含任何公司印章/执照数据**。
真实数据存放在 Supabase（或本地 CSV），不在版本库中。

## 目录
- `index.html`：系统主界面（看板 / 查询 / 柜子图 / 出入库 / 流转轨迹 / 盘点导出 / 数据管理）
- `backend/`：Flask 后端源码（方案 B 部署到 Render/Railway 时使用）

## 部署
- GitHub Pages：本仓库开启 Pages 后会自动发布 `index.html`
- 后端数据：见 Supabase 集成方案

> 注意：当前 `index.html` 通过 `/api/*` 调用后端。若仅发布到 Pages 而未接入
> 后端或 Supabase，页面加载时数据接口会失败。请按部署方案完成数据层接入。
