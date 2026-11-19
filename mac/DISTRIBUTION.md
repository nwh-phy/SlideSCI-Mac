# SlideSCI Mac 0.2.0 安装与发行说明

此 ZIP 面向 Apple Silicon Mac，内置 Node.js 24.22.3、生产依赖及 Swift 原生辅助程序。使用者无需安装 Node.js、Python 或 Xcode。原生程序按 macOS 14 最低版本构建；较早系统及 Intel Mac 安装包未验证。PowerPoint 需要 PowerPointApi 1.10（Mac 16.105 起）及 SharedRuntime 1.1；Office.js 首次加载需要访问微软 CDN。

这是 100% 本地托管加载项：WPS 页面访问 `http://127.0.0.1:3000`，PowerPoint 页面访问 `https://localhost:3443`，不监听局域网。`Install.command` 会自动安装并启用 macOS LaunchAgent（`~/Library/LaunchAgents/com.slidesci.mac.service.plist`），随系统登录静默运行，重启电脑后无需再手动运行 `Start.command` 或保留终端窗口。

核心科研组图功能（图片自动排列、六种对齐、间距与均匀分布、位置与宽高复制粘贴、全选文本框、添加/更新图片标签、添加上/下标题）全部在 PowerPoint 中纯 Office.js 本地直跑，零剪贴板触碰、零 AppleScript、无需系统自动化权限；剪贴板自动监听默认关闭（可在「设置与帮助 → 偏好设置与帮助」手动开启），不再有 6 秒一次的后台整页 PPTX 导出。只有使用高级原生功能（记录原图路径、16 位 TIFF、伪彩、放大图、比例尺、公式/Markdown、素材库、高清导出）时才按需调用本地服务。

## 安装与启用

1. 先安装并打开需要使用的 WPS 或 PowerPoint 一次，再保存现有工作。安装器只注册已经初始化容器的宿主。
2. 解压 ZIP，运行 `Install.command`。程序安装到 `~/Library/Application Support/SlideSCI-Mac`，同时写入登录自启配置并启动本机后台服务。已有旧版时先运行旧版本的 `Uninstall.command`，素材和设置会保留。
3. 如需手动启停服务，可运行 `Start.command` 或 `Stop.command`。若提示 3000/3443 端口占用，先结束占用端口的开发服务。
4. WPS：保存工作并重开 WPS，确认信任 `slidesci-mac-dev`，从顶部 SlideSCI 分组按钮直接选择功能。
5. PowerPoint：PowerPoint 只通过受 macOS 信任的 HTTPS 加载 SlideSCI。安装时会为本机生成独立的 `localhost` 证书和私钥（有效期 825 天），保存在 `~/Library/Application Support/SlideSCI-Mac-Installer/tls/`，私钥只保存在本机。安装器会询问是否信任该证书（仅用于 `localhost` 的 SSL，不创建 CA）；同意后输入登录密码即完成。证书剩余 30 天内可运行 `RenewCertificate.command` 换新。
6. 保存工作并重开 PowerPoint，从“开始 → 加载项 → 开发人员加载项”双击 `SlideSCI Mac`。

ZIP 和原生辅助程序尚无 Developer ID 签名及 Apple 公证。Gatekeeper 可能阻止首次打开；确认来源后在系统设置“隐私与安全性”中允许该文件即可。

服务启动采用完全惰性策略，不会主动探测 PowerPoint/WPS 容器目录或弹出权限窗；仅当你首次触发需要文件/自动化的高级功能时，macOS 才可能按需请求相应权限。本版本已移除 AI 助手功能。

## 停止与卸载

- `Stop.command` 会通过 `launchctl bootout` 停止本机自启服务，并清理过期进程记录，不结束任何其他程序。
- `Uninstall.command` 先停止并删除 LaunchAgent 自启配置，再恢复安装前的宿主加载项注册。其他加载项保留。
- 卸载将程序目录移入相邻 `.uninstalled-时间` 备份目录，保留素材、设置与安装器备份，并移除安装器添加的证书信任。
- 业务数据沿用 `~/Library/Application Support/SlideSCI Mac`；安装器状态在 `~/Library/Application Support/SlideSCI-Mac-Installer`。卸载不会移除此前独立安装的开发版。

## 验证与限制

WPS 已在独立文稿中验证局部文字格式、物理比例尺、局部放大图、Markdown 原生公式及原生公式保存重开后编辑。局部颜色复制保留其余字符颜色与 22 磅字号；保存重开后比例尺长度仍为原图宽度的 20%，放大框、放大图及两条辅助线记录保留。原生公式在重开后通过 WPS 公式编辑器改写分子并保存，Office Math 结构保留。

PowerPoint 在不含旧开发加载项引用的新建文稿中完成 19 项核心实机检查，全部通过。覆盖局部文字格式、图片来源与伪彩、物理比例尺、放大图、Markdown 与原生公式、原图提取、PNG/JPEG/BMP/PDF 导出及可编辑素材保存和复用。保存重开后颜色、22 磅字号、20% 比例尺、放大图和两条辅助线记录保留；原生公式重开后经 PowerPoint 公式编辑器改写并保存，仍保留两处 Office Math。修复了 PowerPoint 粘贴重复对象时导致的 12 磅位置偏移。

隔离安装验证覆盖宿主注册与备份、生产依赖、原生辅助程序、实际内容转换、端口冲突、PID 归属、停止服务及卸载恢复。当前发行的检查结果见同一 Release 中的 `INSTALL-VERIFICATION.json`。全新 macOS 账户的 Gatekeeper、权限提示，以及安装器触发的证书信任授权窗口尚未实机验证；安装包没有 Apple 公证。

服务启动时清理宿主暂存目录中超过 24 小时的中间文件（视频、音频文件保留），导出完成后立即删除中间 PPTX 与 PDF。侧栏与服务断开后自动重连；服务重启后，侧栏自动恢复粘贴监听，不需要重新加载插件。侧栏关闭或宿主退出 60 秒后，剪贴板监听自动停止。

不将浏览器预览、HTTP 200 或构建成功视为宿主功能通过。字体缺失可能影响显示；源图片路径在另一台机器上未必存在；多选顺序取决于宿主返回顺序。外部 AI 请求需要用户配置且可能发送所选文稿内容。

## 问题反馈

Mac 版问题请在 [SlideSCI-Mac Issues](https://github.com/nwh-phy/SlideSCI-Mac/issues) 反馈，附上 macOS、PowerPoint 或 WPS 版本，以及安装器状态目录中的 `service.log`。原版 Windows 插件的作者不维护 Mac 版。

## 来源与校验

`BUILD-INFO.json` 记录版本、内置 Node 版本及归档 SHA-256、构建时间、Git 基线和源码快照指纹。该 Git 基线不代表快照中的未提交改动已公开提交。`SOURCE-MANIFEST.json` 对随包源码逐文件记录 SHA-256；`mac/` 和 `SlideSCI/latex-converter/` 包含本发行运行所需的源码及公式资源。`LICENSE` 为 AGPL-3.0；Node、生产依赖及 Phosphor 图标保留各自许可证。

同目录 `.zip.sha256` 用于核对下载文件。仅执行 `shasum -a 256 -c 文件名.zip.sha256` 不会安装软件。

开发者构建：在有 Node.js、npm 与 Xcode Command Line Tools 的 Apple Silicon Mac 上安装 `mac/` 的锁定依赖后，执行 `node mac/scripts/distribution.mjs build`。构建器校验官方 Node 归档，安装锁定生产依赖，编译原生程序并生成 ZIP 与 SHA-256。随后执行 `node mac/scripts/verify-package.mjs 安装包.zip --report INSTALL-VERIFICATION.json`，在隔离目录中检查包内容、内置运行时、原生程序、公式转换、安装注册、过期服务记录、证书换新和卸载恢复；3000/3443 端口空闲时还会检查服务启动与停止。
