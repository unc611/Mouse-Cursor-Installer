# 鼠标光标批量安装工具（Mouse-Cursor-Installer）

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)

**简体中文** | [**English**](./README_en.md)

一个简单但功能强大的 Windows 鼠标光标方案批量安装工具，支持智能识别、自动匹配和一键安装。

## ✨ 特色功能

### 🎯 智能光标类型识别

- **多语言关键词识别：** 支持中文、英文、日文光标文件名识别
- **智能前后缀处理：** 自动识别并去除文件名中的公共前缀和后缀
- **数字序号匹配：** 支持基于数字序号的光标类型推断
- **备用匹配机制：** 当关键词匹配失败时，使用文件序号作为备用方案

### 📦 批量安装

- **通用方法：** 不再需要`.inf`文件，适用于各种光标方案
- **自动设置方案名：** 自动以文件夹名作为方案名，快速安装多个光标主题
- **灵活的目录结构：** 支持递归扫描子文件夹

### 🔄 自动替代机制

- **Wait ↔ AppStarting 互相替代：** 当其中一个缺失时自动使用另一个
- **Hand/Pin/Person 优先级替代：** 按优先级顺序自动填补缺失的光标类型

### 📊 完整的安装报告

- **数量异常检测：** 识别非常见数量（5/10/15/17）的光标方案
- **未匹配文件检测：** 发现可能存在命名问题的光标文件
- **详细统计报告：** 提供安装成功数量、处理文件数量等详细信息

> ⭐ 如果这个项目对您有帮助，请给它一个星标！

## 🚀 使用方法

1. **下载：** 前往 [Releases](https://github.com/unc611/Mouse-Cursor-Installer/releases) 页面，下载最新的 `Mouse-Cursor-Installer.exe` 文件。

2. **文件夹结构：**
   
   - 在任意位置，创建你的光标文件夹。**每个主题一个文件夹，文件夹名称将成为鼠标方案的名称。** 例如：
   
   ```
   MyCursors/
   ├── 方案名1/
   │   ├── cursor1.ani
   │   ├── cursor2.ani
   │   └── ...
   ├── 方案名2/
   │   ├── Arrow.cur
   │   ├── help.cur
   │   └── ...
   ├── 正常选择.cur  # 也支持根目录下的光标文件
   ├── 帮助选择.cur
   └── ...
   ```

3. **运行工具：**
   
   - 双击运行 `Mouse-Cursor-Installer.exe`，程序会自动请求管理员权限。
   
   - 根据菜单提示进行操作，只需几秒就能完成！

4. **应用方案：** 安装完成后，工具会询问是否立即打开鼠标属性面板。你可以在「指针」选项卡下的「方案」下拉菜单中找到并应用你刚刚安装的所有主题。

## 🔧 构建指南
   
   如果你想从源码运行或自行修改，请遵循以下步骤：

1. **环境要求：** Windows 10/11，PowerShell 5.1 或更高版本。

2. **直接运行脚本：**
      
      - 克隆仓库到本地。
      
      - 在 PowerShell 中，执行以下命令（需要管理员权限）：
  
    ```powershell
      Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
      ./mouse-cursor-installer.ps1
    ```

3. **打包成 `.exe`**：
      
      - 本项目使用 [PS2EXE](https://github.com/MScholtes/PS2EXE) 工具进行打包。
      
      - 安装 PS2EXE 模块：
  
    ```powershell
      Install-Module ps2exe
    ```

      - 执行以下命令进行打包：

    ```powershell
      ps2exe .\mouse-cursor-installer.ps1 .\mouse-cursor-installer.exe
    ```

## 🐛 故障排除

### 常见问题

**Q: 双击运行后自动提权失败**

```
A: 尝试手动右键以管理员身份运行
```

**Q: 光标文件无法识别或识别错误**

```
A: 检查文件名是否过于不规范
```

**Q: 安装后在控制面板中找不到新方案**

```
A: 尝试重新打开"鼠标属性"窗口
```

**Q: 方案名称不符合预期**

```
A: 工具会自动以文件夹名作为方案名称，请修改文件夹名
```

### 调试模式

在主菜单输入-debug来启用调试模式，可以获得详细的处理日志。

日志将保存到程序目录下的mouse-cursor.log

## 📄 编写

- 59%的代码由Deepseek R1，30%的代码由Claude sonnet 4，10%的代码由Gemini 2.5 pro，剩下1%由本人编写。
- 本人只负责调试和处理一些逻辑问题。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！
