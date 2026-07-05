# SUAT-cats

## 写在前面
### 为什么制作这个网站
记录SUAT的猫猫，给SUAT的猫猫制作一个花名册。

### 我也想上传图片/修改文案/完善猫猫信息等
VX搜索“无蔗糖”并添加好友，因为制作网站时间仓促、人数不多，网站难免有些瑕疵。欢迎屏幕前的各位帮编者完善猫猫信息！！！

目前不接受 Pull requests

### 这个网站安全吗 & 为什么用这个网站
本网站使用了GitHub拖管服务，GitHub是地球上最大的代码分享网站，基本每个学计算机的都会使用到，100%安全。

GitHub可以提供免费的网站托管服务，也就是说在这上面建一些网站是免费的，不需要花钱买服务器。

买服务器要考虑到防攻击、维护、流量、费用等等一系列问题，服务器提供商可能刚开始卖的比较便宜，到后面就越来越贵，每年维护服务器的成本都可以绝育5只公猫，或者绝育+引产1只母猫。

在GitHub这个网站上，任何的代码都是开源的，包括这个网站代码也是开源的，这意味着可以有更多的人贡献到网站的制作上来。

这些代码不仅有显示猫猫的前端，也有添加、修改猫猫的可视化后端，维护者可以轻松修改里面的猫猫信息，不用了解头疼的代码。

等编者（目前是"无蔗糖"）毕业了，我也会在这个文档后面写一些教程，教维护者如何添加猫猫、修改猫猫信息等等。

# 下面的不必看啦
## 代码使用教程（完善中）
### 项目地址
[点我跳转](https://github.com/Tommy-150/SUAT-cats)

### 目录核心文件
ban_words.json - 禁用一些敏感词

cats.json - 猫猫所有信息（重要）

index_mobile.html - 手机端前段

index_pc.html - 电脑端前段

index.html - 把用户导向pc或电脑

tag_color.json - 设置tag的颜色

preview_web.py - 预览网页

cats_manager.py - 控制台（重要）

### 使用方法
大部分功能操作都集成到了 cats_manager.py 里面

不出意外千万不要碰 cats_manager.py 以外的任何文件！！！

运行需要这些依赖：
```bash
pip install Pillow openpyxl pywebview
```

## 打包版 EXE 使用说明

### 是什么
`SUATCatManager.exe` 是一个开箱即用的独立可执行文件，放在项目文件夹里双击就能运行管理器，不需要安装 Python 或任何依赖。适合完全不懂代码的维护者。

### 打包技术原理
使用 **PyInstaller** 将 Python 脚本和所有依赖打包成单个 .exe：

1. **静态分析**：PyInstaller 扫描 `app.py` 的 import 链，把 `webview`、`openpyxl`、`Pillow`、`numpy` 等所有用到的库都收集起来
2. **字节码打包**：所有 .py 文件编译成 .pyc 字节码，和 .dll 一起压入可执行文件
3. **自解压引导**：exe 启动时，bootloader 在系统临时目录 `%TEMP%` 下创建 `_MEIxxxxx` 文件夹，把打包的所有文件解压进去，然后用内嵌的 Python 解释器运行
4. **窗口模式**：`--windowed` 参数让 exe 不弹命令行黑窗，直接显示 pywebview 图形窗口

### exe 运行需要什么
| 依赖 | 说明 |
|---|---|
| Windows 10/11 | pywebview 调用系统内置的 Edge WebView2 渲染页面，Win10+ 自带 |
| 项目文件夹 | exe 必须和 `classified/`、`统计信息.xlsx`、`cats.json` 等放在同一目录，程序读写都在 exe 所在目录 |
| 不需要的 | 不需要 Python、不需要 pip、不需要装任何东西 |

### 如何重新打包
如果改了 `app.py` 或 `app_manager.html`，重新生成 exe：
```bash
pip install pyinstaller pywebview
pyinstaller --onefile --windowed --name SUATCatManager app.py
# exe 生成在 dist/ 目录，复制到项目根目录即可
```

## 特别鸣谢
deepseek-v4-pro完成了大部分代码

qwen3.7-max完成了demo1-4的部分，为我们的网页确定了主题

GLM-5.2完成了cat_manager.py的部分内容