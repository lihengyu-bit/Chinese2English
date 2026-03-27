# PPT 中英翻译与自动排版

这是一个基于 `Streamlit` 的单页面 Web 工具，用于在浏览器中上传 `.pptx`、调用 Gemini API 完成中文到英文翻译，并生成新的 PPT 文件下载。

## 功能说明

- 输入用户自己的 Gemini API Key
- 拖拽或点击上传 `.pptx`
- 逐页显示处理进度
- 翻译完成后直接下载新的 `.pptx`
- 递归处理组合图形中的文本
- 仅翻译形状和表格中的文字，保护图片、图表、背景和整体结构
- 译文较长时，对文本框高度、表格行高以及下方重叠元素做保守式调整

## 给最终使用者的说明

- 浏览器使用者不需要安装任何 Python 依赖
- 只有部署这套网页的服务器或管理员机器，需要先安装一次依赖
- 依赖安装完成后，其他同事只需要打开浏览器访问网页地址

## Windows 一键启动

如果是在 Windows 服务器或办公电脑上部署，优先直接双击：

```bash
start_webapp.bat
```

这个脚本会自动：

- 创建本地虚拟环境 `.venv`
- 升级 `pip`
- 安装 `requirements.txt`
- 启动 Streamlit 网页服务

默认地址：

- 本机访问：`http://127.0.0.1:8501`
- 局域网访问：`http://服务器IP:8501`

## 手动运行方式

1. 安装依赖

```bash
pip install -r requirements.txt
```

2. 启动网页

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

3. 在浏览器打开服务器地址

## Docker 部署

如果公司内部更适合标准化部署，可以直接使用 Docker：

```bash
docker build -t ppt-translator .
docker run --rm -p 8501:8501 ppt-translator
```

## 实现要点

- Gemini 接口：`https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent`
- 默认模型：`gemini-2.5-flash`
- 跳过纯数字和不含中文的文本，减少 Token 消耗
- 优先通过 `python-pptx` 递归遍历 `Shapes`、`Tables` 和组合图形内部元素
- 如果当前 Windows 电脑已安装 Microsoft PowerPoint，也可以自动切换到本机 PowerPoint 引擎
- 尽量保留字体颜色、加粗、斜体等原始样式

## 注意事项

- 该工具不会保存 API Key 到磁盘，刷新页面后需要重新输入
- 对复杂混排、极端拥挤版式和高度定制动画，自动排版只能做保守微调，建议导出后再人工抽查关键页面
- 如果部署机器无法联网，需要由管理员提前准备好 Python 包镜像或离线 wheel 文件

## 固定网址部署建议

如果你要的是“点开一个固定网址就能用”，不要继续使用本机 `http://127.0.0.1:8501` 这种临时地址。

推荐方式：

- 部署到长期在线的云主机或容器平台
- 使用当前项目自带的 `Dockerfile`
- 云端安装 `python-pptx` 后，服务会直接走标准服务端引擎，不依赖本机 Microsoft PowerPoint

当前仓库已经补好了适合固定网站部署的文件：

- `requirements.txt`：已包含 `python-pptx`
- `Dockerfile`：已支持平台注入的 `PORT`
- `render.yaml`：可直接用于 Render 这类平台

这样部署完成后，平台会分配一个固定网址；如果你再绑定自己的域名，网址也可以长期不变。
