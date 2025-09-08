# PDF转Word工具

这是一个基于Python的PDF转Word转换工具，具有以下功能：

1. 使用pdf2docx库实现PDF到Word的转换
2. 提供Web界面供用户上传PDF文件
3. 支持用户选择转换的页码范围
4. 可以打包成Windows可执行文件

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行应用

```bash
python src/app.py
```

访问 http://localhost:5000 使用应用。

## 打包为Windows可执行文件

```bash
pyinstaller --onefile --add-data "src/templates;templates" --add-data "src/static;static" src/app.py
```

打包后的可执行文件将在dist目录中生成。