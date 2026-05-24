# 作业文档转手写体工作台

这是我的个人私有工具，用来把作业、讲义、笔记、论文草稿等文档整理成可预览、可导出的手写体结果。项目已经从原来的公开手写生成站改造成偏本地/私有部署的工作台，所以这里不再保留公开演示站、宣传截图、赞助入口、社群链接和公开 SEO 文案。

## 当前功能

- 支持直接粘贴正文，或上传 PDF、Word、Markdown、TXT、RTF 作为正文来源。
- 上传文档后会抽取内容并规范化为 Markdown，数学公式会尽量整理成可读形式。
- PDF 抽取可接入 MinerU 服务，用于把扫描或复杂排版文档转成 Markdown。
- 支持生成标准 Word 校对稿，便于先检查正文和公式。
- 支持手写体预览，并导出 PDF 或 Word 文件。
- 保留字体、背景图、边距、字号、行距、扰动、墨色、涂改等手写渲染参数。
- 后端使用任务队列和 WebSocket/轮询进度，避免长文档生成时前端一直卡住。

## 目录结构

```text
backend/          FastAPI 后端、文档抽取、公式整理、手写渲染
frontend/         Vue 前端工作台
ttf_files/        本地手写字体
mysql/            可选数据库初始化文件
docker-compose.yml 本地私有部署配置
```

## 本地开发

Python 环境请使用 Conda 或 uv 管理，不要使用 `sudo pip`，也不要混用 Homebrew Python、系统 Python 和项目环境。

```bash
conda create -n handwriting-web python=3.11
conda activate handwriting-web
pip install -r backend/requirements.txt
```

启动后端：

```bash
cd backend
python app.py
```

启动前端：

```bash
cd frontend
npm install
npm run serve
```

开发访问地址：

```text
http://localhost:8080
```

后端默认监听：

```text
http://127.0.0.1:5005
```

## PDF 抽取配置

如果要使用 MinerU 抽取 PDF，需要在后端运行环境中配置：

```bash
export MINERU_BASE_URL="https://你的-mineru-服务地址"
export MINERU_API_TOKEN="你的-token"
export MINERU_PUBLIC_BASE_URL="http://你的后端可公网访问地址"
```

本地只处理 Markdown、TXT、Word 时可以不配置 MinerU；上传 PDF 时如果缺少配置，后端会返回明确错误。

## 私有部署

Docker Compose 以本地构建为主：

```bash
docker compose up --build -d
```

默认端口：

- 前端：`2345`
- 后端：`127.0.0.1:5005`

字体放在 `ttf_files/`，Compose 会挂载到后端容器。

## 验证命令

后端重点测试：

```bash
python -m unittest backend.tests.test_unified_handwriting_pipeline
```

前端构建：

```bash
cd frontend
npm run build
```

## 维护备注

- 这个仓库按个人私有项目维护，不再面向公开站点运营。
- 不提交运行日志、缓存、临时文件、构建产物和本地代理配置。
- 提交前检查 `git status`，确认只包含本次需要的代码和文档改动。
