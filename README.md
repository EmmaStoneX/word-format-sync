# Word 格式识别与套用工具

一个基于 Python、PySide6 与 Microsoft Word COM 的桌面工具，用来从模板文档提取格式规则，并将这些规则应用到目标 Word 文档。

## 下载 exe

已提供 exe 版本，发布后可直接前往 GitHub 的 Tags / Releases 页面下载对应版本。

## 致谢

感谢产品经理 wublub，感谢大模型 gpt-5.4。

## 功能简介

- 支持拖拽或选择 `doc` / `docx` 文档
- 支持两种识别方式：`模板页抽取`、`范围识别`
- 支持按页码范围将格式应用到目标文档
- 目标文档上传后自动加载完整检查项目录
- 支持检查项拖拽排序
- 支持标题搜索与快速定位
- 支持标题快速移动，可将单个顺序或一段顺序移动到指定顺序下方
- 支持标题编号预览与重编号
- 支持导入 / 导出格式参数 JSON
- 输出到新文件，不直接覆盖原始文档

## 运行环境

- Windows 10 / 11
- 已安装桌面版 Microsoft Word
- Python 3.10 及以上

> 本项目依赖 Word COM 自动化，不适用于未安装桌面版 Word 的环境。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动命令

```bash
python app.py
```

## 使用步骤

1. 选择或拖入模板文档
2. 选择识别模式并点击“识别格式”
3. 按需要调整分类格式参数
4. 选择或拖入目标文档
5. 检查右侧检查项、搜索标题或调整标题顺序
6. 设置输出文件名
7. 点击“应用到目标文档”生成结果

## 项目结构

```text
word格式/
├─ app.py
├─ README.md
├─ LICENSE
├─ requirements.txt
├─ models/
├─ services/
├─ ui/
└─ utils/
```

## 打包

如需本地打包，可使用：

```bash
pyinstaller -F -w app.py -n "Word格式识别与套用工具"
```

## 许可证

本项目使用 MIT License，详见 `LICENSE`。
