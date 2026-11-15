# Lean Design Tools

一个独立、轻量的 Windows PowerPoint 设计任务窗格。界面与操作逻辑按可观察行为 clean-room 重写，不包含 iSlide 的代码、图标或资源。

许可证：MIT。

当前实现：

- 对齐与等距分布
- 等宽、等高、等大
- 贴合幻灯片边缘、居中和铺满
- 全选、组合、取消组合、图层顺序、选择窗格
- 联合、组合、拆分、相交和剪除形状
- 跨幻灯片复制/粘贴位置、宽度、高度和大小，四类缓存互不覆盖
- 四向吸附、旋转和翻转
- 单字拆分/合并、按段拆分/段落合并文本框
- 文本框自适应、自动换行和四边距
- 对象宽度、高度的厘米输入与选中状态回显，图片和普通形状均可使用
- 选中状态下回显文本框模式、自动换行和边距
- 文本框区红色状态控件、清空文本、带 `cm` 单位的自绘平面边距输入和逐项悬浮说明；方向标记采用“边界线＋箭头”，无文本选择时使用 iSlide 式浅灰状态
- iSlide 风格的“布局 / 形状 / 对象”扁平页签，选中项使用珊瑚红文字和下划线；每页工具区全部展开、无折叠动画
- 对齐、大小和参考线布局分区之间使用 1 px 浅灰横线
- 17×15、1 px 珊瑚红 clean-room 线性图标，按 iSlide 的可观察几何语义重绘，统一圆角端点、悬浮与灰色不可用状态
- 所有工具（包括灰色暂不可用状态）均可显示悬浮说明

## 构建

```powershell
.\build.ps1
```

## 安装

普通用户请从 GitHub Releases 下载 `LeanDesignTools-Setup-x64-版本号.exe`，保存并关闭 PowerPoint 后双击安装。安装程序会检查 64 位 Office 和 .NET Framework 4.8，完成 COM 注册并添加卸载入口。

开发环境也可以运行：

保存并关闭 PowerPoint，然后运行：

```powershell
.\install.ps1
```

安装脚本需要管理员权限，用 64 位 .NET Framework RegAsm 注册 COM 类；PowerPoint 加载项开关仍只写入当前用户。卸载使用 `.\uninstall.ps1`。

## 构建安装包

安装 Inno Setup 6 后运行：

```powershell
.\build-installer.ps1
```

生成文件位于 `dist\LeanDesignTools-Setup-x64-版本号.exe`。

## 限制

- 当前只构建 x64，与本机 64 位 Office 匹配。
- 已在本机 64 位 PowerPoint LTSC 2024 实测加载、任务窗格显示、等大、文本框自适应和状态回显。
- 构建依赖本机安装的 Office 与 Extensibility PIA。
- 图标为 clean-room 自绘线框图标，不复用 iSlide 资源；细节不会像素级复制其受版权保护的素材。

完整的自动化与真实 PowerPoint 验证结果见 [VERIFICATION.md](VERIFICATION.md)。
