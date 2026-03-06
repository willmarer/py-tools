# Offline PPT Translator

一个离线、轻量、可用的 PPT 文本替换翻译器（V1）。

## 功能
- GUI 选择输入文件夹 / 输出文件夹
- 扫描 `.pptx` 和 `.ppt`
- 批量处理 `.pptx`
- `.ppt` 仅扫描提示，V1 不处理
- 快速测试翻译
- 日志输出
- 词典独立 JSON 文件，方便扩展

## 安装依赖
```bash
pip install -r requirements.txt

## 打包exe
```bash
pyinstaller --noconsole --onefile --add-data "phrases.json;." --add-data "lexicon.json;." app.py
