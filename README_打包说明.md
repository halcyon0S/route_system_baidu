# 打包为EXE文件说明

## 方法一：使用批处理文件（推荐）

### 标准打包（平衡体积和兼容性）

1. 确保已安装Python和pip
2. 双击运行 `build.bat`
3. 打包完成后，exe文件位于 `dist\route_system_baidu.exe`
4. 文件大小通常为 50-80 MB

### 高级优化打包（最小体积）

1. 双击运行 `build_optimized.bat`
2. 会进行深度优化，进一步减小体积
3. 文件大小通常为 40-60 MB
4. **注意**：可能需要更长的打包时间

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
    --optimize=2 ^
    --strip ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import=pandas._libs.tslibs.timedeltas ^
    --hidden-import=openpyxl ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=tkinter ^
    app.py
```

## 体积优化说明

### 已实施的优化措施：

1. **排除不必要的模块**：
   - 测试框架（pytest, unittest）
   - 开发工具（jupyter, IPython）
   - 图形库（matplotlib, tkinter, PyQt）
   - 科学计算库（scipy）
   - 文档工具（sphinx, pydoc）

2. **Pandas优化**：
   - 只包含必需的pandas组件
   - 排除测试和绘图功能
   - 排除不必要的IO模块

3. **编译优化**：
   - 启用 `strip=True` 移除调试信息
   - 启用 `optimize=2` Python优化级别
   - 启用UPX压缩（如果可用）

4. **精简导入**：
   - 只包含实际使用的hiddenimports
   - 排除未使用的pandas子模块

### 进一步减小体积的方法：

1. **安装UPX压缩工具**：
   - 下载：https://upx.github.io/
   - 添加到系统PATH
   - 重新打包会自动使用UPX压缩

2. **使用虚拟环境**：
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   # 然后打包
   ```

3. **检查依赖**：
   - 只安装必需的包
   - 避免安装开发依赖

## 运行EXE

1. 双击 `dist\route_system_baidu.exe` 运行
2. 程序会在后台启动Flask服务器
3. 打开浏览器访问：http://127.0.0.1:5005

## 注意事项

- 首次运行可能需要几秒钟启动时间
- 如果杀毒软件报毒，这是误报，可以添加信任
- 确保防火墙允许程序访问网络（用于调用百度地图API）
- 如果需要显示控制台窗口查看日志，可以修改 `build_exe.spec` 中的 `console=True`
- 如果遇到"找不到模块"错误，检查 `hiddenimports` 列表

## 文件说明

- `build_exe.spec`: PyInstaller配置文件（已优化）
- `build.bat`: Windows批处理打包脚本（标准打包）
- `build_optimized.bat`: 高级优化打包脚本（最小体积）
- `requirements.txt`: Python依赖列表

## 常见问题

### Q: exe文件太大怎么办？
A: 
1. 使用 `build_optimized.bat` 进行高级优化
2. 安装UPX压缩工具
3. 检查是否有不必要的依赖

### Q: 打包后运行报错"找不到模块"？
A: 
1. 检查 `build_exe.spec` 中的 `hiddenimports` 列表
2. 添加缺失的模块到 `hiddenimports`
3. 重新打包

### Q: 如何进一步减小体积？
A: 
1. 使用虚拟环境，只安装必需包
2. 安装UPX并启用压缩
3. 考虑使用 `--onedir` 模式（文件夹模式）而不是 `--onefile`

