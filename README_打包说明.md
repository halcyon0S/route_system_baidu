# 打包为EXE文件说明

## 方法一：使用批处理文件（推荐）

1. 确保已安装Python和pip
2. 双击运行 `build.bat`
3. 打包完成后，exe文件位于 `dist\route_system_baidu.exe`

## 方法二：手动打包

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 使用PyInstaller打包

```bash
pyinstaller build_exe.spec --clean
```

### 3. 或者直接使用命令行打包

```bash
pyinstaller --name=route_system_baidu ^
    --onefile ^
    --windowed ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import=pandas ^
    --hidden-import=flask ^
    --hidden-import=openpyxl ^
    app.py
```

## 运行EXE

1. 双击 `dist\route_system_baidu.exe` 运行
2. 程序会在后台启动Flask服务器
3. 打开浏览器访问：http://127.0.0.1:5004

## 注意事项

- 首次运行可能需要几秒钟启动时间
- 如果杀毒软件报毒，这是误报，可以添加信任
- 确保防火墙允许程序访问网络（用于调用百度地图API）
- 如果需要显示控制台窗口查看日志，可以修改 `build_exe.spec` 中的 `console=True`

## 文件说明

- `build_exe.spec`: PyInstaller配置文件
- `build.bat`: Windows批处理打包脚本
- `requirements.txt`: Python依赖列表

