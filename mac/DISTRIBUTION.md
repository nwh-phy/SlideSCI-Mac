# SlideSCI Mac 0.1.2 安装与发行说明

此 ZIP 面向 Apple Silicon Mac，内置 Node.js 24.22.3、生产依赖及 Swift 原生辅助程序。使用者无需安装 Node.js、Python 或 Xcode。原生程序按 macOS 14 最低版本构建；较早系统及 Intel Mac 安装包未验证。实际验证环境为 macOS 27.2、WPS 12.1.29166、PowerPoint 16.115。PowerPoint 需要 PowerPointApi 1.10（Mac 16.105 起）及 SharedRuntime 1.1；Office.js 首次加载需要访问微软 CDN。

这是本地加载项：WPS 页面访问 `http://127.0.0.1:3000`，PowerPoint 页面访问 `https://localhost:3443`。使用期间必须保持本机服务运行；关闭终端窗口不会自动结束已启动的服务。服务不监听局域网，也不随开机自动启动：电脑重启或注销后，先运行 `Start.command`，再打开 WPS 或 PowerPoint。服务记录在重启后过期时，`Start.command` 会自动清除旧记录并启动，不会结束其他程序。

## 安装与启用

1. 先安装并打开需要使用的 WPS 或 PowerPoint 一次，再保存现有工作。安装器只注册已经初始化容器的宿主。
2. 解压 ZIP，运行 `Install.command`。程序安装到 `~/Library/Application Support/SlideSCI-Mac`。安装器保存现有 SlideSCI 注册的备份，并保留其他加载项。已有本安装器版本时，先运行旧版本的 `Uninstall.command`，再安装新版本；素材和设置会保留。
3. 从安装目录运行 `Start.command`。若提示端口占用，先结束占用 3000/3443 的开发服务；安装器不会结束不属于自己的进程。
4. WPS：保存工作并重开 WPS，确认信任 `slidesci-mac-dev`，从顶部 SlideSCI 的排列与尺寸、图片处理、科研标注、内容插入、素材与导出、助手与设置分组直接选择功能。
5. PowerPoint：PowerPoint 只通过受 macOS 信任的 HTTPS 加载 SlideSCI。安装时会为本机生成独立的 `localhost` 证书和私钥（有效期 825 天，即 macOS 允许的上限），保存在 `~/Library/Application Support/SlideSCI-Mac-Installer/tls/`，私钥只保存在本机，不随发行包分发。安装器随后询问是否信任该证书；同意后 macOS 弹出授权窗口，输入登录密码即完成。信任范围仅限这一张证书、仅用于 `localhost` 的 SSL 连接，安装器不创建证书颁发机构（CA）。暂不同意也可完成安装，WPS 照常可用；之后运行 `Start.command` 会再次询问，也可用“钥匙串访问”手动导入 `localhost.crt` 并将 SSL 设为始终信任。批量部署可在命令后加 `--trust-certificate yes` 或 `--trust-certificate no` 预先作答。证书剩余 30 天内，`Start.command` 会提醒；运行安装目录中的 `RenewCertificate.command` 可提前换新，已过期时 `Start.command` 会自动换新。换新时会移除旧证书的信任、询问是否信任新证书，旧证书文件移入 `tls/retired-时间` 目录保留。
6. 保存工作并重开 PowerPoint，从“开始 → 加载项 → 开发人员加载项”双击 `SlideSCI Mac`。加载后从顶部 SlideSCI 分组按钮或菜单直接选择功能。需要参数的功能会打开对应表单或预览，无参数快捷操作直接执行；侧栏通过“功能”选择器切换，一次只显示一个工具。以任务窗格显示“已连接 Microsoft PowerPoint”为启用依据。安装器不会更改隐私设置或账户。

ZIP 和原生辅助程序尚无 Developer ID 签名及 Apple 公证。Gatekeeper 可能阻止首次打开；需要使用者检查来源后，在系统设置“隐私与安全性”中允许该下载文件。不要关闭 Gatekeeper 或全局删除隔离属性。不同 macOS 版本的提示可能不同。

按需使用的文件选择、Office 自动化、剪贴板监测和 TIFF 粘贴拦截，可能要求 macOS 文件、自动化、辅助功能或输入监控权限。服务需要在 PowerPoint、WPS 各自的数据目录中暂存图片和导出文件，首次使用时 macOS 可能提示终端“访问其他 App 的数据”，需要允许；拒绝后相关操作会给出提示，允许后重新运行 `Start.command`。请仅为实际运行安装服务的终端/辅助程序授予所需权限。AI 服务需要自行配置供应商地址、模型与密钥；真实供应商请求尚未验证。

## 停止与卸载

- `Stop.command` 只结束本安装器记录并核对过的 SlideSCI 服务进程组；记录的进程已退出或进程号已被其他程序使用时，只清除过期记录，不结束任何进程。
- `Uninstall.command` 先停止本服务，再移除本次注册并恢复安装前的 SlideSCI 注册。其他加载项保留。注册被另行修改时会拒绝覆盖，保留现场与备份。
- 卸载将程序目录移入相邻 `.uninstalled-时间` 备份目录，保留素材、设置、AI 密钥与安装器备份；安装器添加的证书信任及钥匙串中的证书副本会被移除（macOS 可能弹出授权窗口），证书文件保留，重装时再次询问。需要彻底删除备份或 AI 密钥时，由使用者单独处理。
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
