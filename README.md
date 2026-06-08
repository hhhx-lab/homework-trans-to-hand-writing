# 作业文档转手写体工作台

这是个实用工具，用来把作业、讲义、笔记、论文草稿等文档整理成可预览、可导出的手写体结果，尤其针对数学公式，图表等等有非常好的导出效果。

## 当前功能

- 支持直接粘贴正文，或上传 PDF、Word、Markdown、TXT、RTF 作为正文来源。
- 上传文档后会抽取内容并规范化为 Markdown，数学公式会尽量整理成可读形式。
- PDF 抽取可接入 MinerU 服务，用于把扫描或复杂排版文档转成 Markdown。
- 支持生成标准 Word 校对稿，便于先检查正文和公式。
- 支持单页预览、全量预览和多页翻页，并导出 PDF 或 Word 文件。
- 保留字体、背景图、边距、字号、行距、扰动、墨色、涂改等手写渲染参数。
- 后端使用任务队列和 WebSocket/轮询进度，避免长文档生成时前端一直卡住。
- 前端采用蓝白纸页风格，左右栏分别独立滚动：左侧设置参数，右侧固定显示预览和翻页控件。

## 目录结构

```text
backend/          FastAPI 后端、文档抽取、公式整理、手写渲染
frontend/         Vue 前端工作台
ttf_files/        本地手写字体
mysql/            可选数据库初始化文件
docker-compose.yml 本地私有部署配置
```

## 前端工作台

默认入口就是可用工作台，不是介绍页。

- 左侧栏：上传或校对正文、调整纸张/字体/边距/扰动参数、导出 PDF/Word。
- 右侧栏：显示手写预览，生成全量预览后可用“首页 / 上一页 / 下一页 / 末页”翻页。
- “预览”只生成单页，适合快速看效果；“全量预览”会请求后端返回所有页面，适合提交前逐页检查。
- PWA 安装提示会延后显示，避免挡住长任务生成和预览翻页。

## 自定义字体

页面上的“上传字体”接受的是 `.ttf` 字体文件，不是照片或手写样张。

- 已有 TTF：直接上传即可用于本次渲染。
- 只有手写照片：需要先用字体制作工具把手写字样生成 `.ttf`，再上传到工作台。
- 只覆盖英文、数字和简单公式时，建议至少准备数字、大小写英文、标点、括号和常用数学符号。
- 中文作业要避免缺字，建议覆盖常用汉字、中文标点、数字、英文、括号和数学符号；只写几个字拍照不足以支撑完整作业渲染。

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

也可以直接在项目根目录一键启动前后端：

```bash
./start-dev.sh
```

脚本会创建或重启 `tmux` 会话 `handwriting-web`，后端窗口运行 `backend/app.py`，前端窗口运行 `npm run serve`。

查看运行日志：

```bash
tmux attach -t handwriting-web
```

停止本项目的前后端：

```bash
./start-dev.sh --stop
```

## PDF 抽取配置

如果要使用 MinerU 抽取 PDF，需要在项目 `.env` 中配置：

```dotenv
MINERU_BASE_URL=https://你的-mineru-服务地址
MINERU_API_TOKEN=你的-token
MINERU_PUBLIC_BASE_URL=http://你的后端可公网访问地址
# 可选：当本机代理/TUN 抢占 Tailscale 路由时，绑定本机 Tailscale 地址
MINERU_BIND_HOST=100.64.0.1
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
