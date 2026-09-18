# 合同排版打印 - 云端部署版

## 项目结构
- `backend/` - Python 后端服务（生成合同+回填附件）
- `frontend/` - 前端插件（多维表格侧边栏）
- `frontend/dist/` - 前端构建产物（部署静态托管）

## 后端部署（Render）
1. 注册 https://render.com （免费）
2. New → Web Service → 连接 GitHub 仓库
3. 配置：
   - Root Directory: `backend`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python contract_server.py`
   - Environment Variables:
     - `FEISHU_PERSONAL_TOKEN`: 你的 PersonalBaseToken
     - `PORT`: 10000（Render 默认）

## 前端部署（Cloudflare Pages）
1. 注册 https://pages.cloudflare.com （免费）
2. Create a project → Direct Upload
3. 上传 `frontend/dist` 文件夹
4. 得到 `https://xxx.pages.dev` 地址

## 飞书配置
1. 多维表格 → 插件 → 自定义插件 → 服务地址填前端 pages.dev 地址
2. 自动化流程 → 发送 HTTP 请求 → URL 填后端 onrender.com/generate
