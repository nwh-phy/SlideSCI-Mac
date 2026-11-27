# SlideSCI Mac 0.2.2 安装与发行说明

此 ZIP 面向 Apple Silicon Mac，内置 Node.js 24.22.3、生产依赖及 Swift 原生辅助程序。使用者无需安装 Node.js、Python 或 Xcode。原生程序按 macOS 14 最低版本构建，安装器和各 `.command` 在更早的系统上直接说明并退出；Intel Mac 未验证。PowerPoint 需要 PowerPointApi 1.10（Mac 16.105 起）及 SharedRuntime 1.1；Office.js 首次加载需要访问微软 CDN。

侧栏页面由本机服务提供：WPS 页面访问 `http://127.0.0.1:3000`，PowerPoint 页面访问 `https://localhost:3443`，不监听局域网。PowerPoint 侧栏另需联网从微软 CDN 加载 Office.js，断网时可能无法打开；WPS 不受影响。`Install.command` 会自动安装并启用 macOS LaunchAgent（`~/Library/LaunchAgents/com.slidesci.mac.service.plist`），随系统登录静默运行，重启电脑后无需再手动运行 `Start.command` 或保留终端窗口。

核心科研组图功能（图片自动排列、六种对齐、间距与均匀分布、位置与宽高复制粘贴、全选文本框、添加/更新图片标签、添加上/下标题）全部在 PowerPoint 中纯 Office.js 本地直跑，零剪贴板触碰、零 AppleScript、无需系统自动化权限；剪贴板自动监听默认关闭（可在「设置与帮助 → 偏好设置与帮助」手动开启），不再有 6 秒一次的后台整页 PPTX 导出。只有使用高级原生功能（记录原图路径、16 位 TIFF、伪彩、放大图、比例尺、公式/Markdown、素材库、高清导出）时才按需调用本地服务。

## 安装与启用

1. 先安装并打开需要使用的 WPS 或 PowerPoint 一次，再保存现有工作并完全退出（`Cmd+Q`）。安装器只注册已经初始化容器的宿主。
2. 解压 ZIP，运行 `Install.command`。程序安装到 `~/Library/Application Support/SlideSCI-Mac`，同时写入登录自启配置并启动本机后台服务。若本机已安装旧版（0.1.x 及以上同路径安装器版本），`Install.command` 支持提示后**原地平滑升级**（自动停止旧服务、备份旧版程序、复用已信任的本地证书免重复输入密码，并在 PowerPoint 未运行且 manifest 有变化时询问是否清理 Wef 加载项缓存；非交互部署可传 `--upgrade yes`、`--clear-wef-cache yes|no`，未作答时不清理）。素材和设置均保留。升级只保留最近一份旧版程序备份（`SlideSCI-Mac.previous-时间`），更早的备份在升级成功后删除。
3. `Start.command` 启动服务并恢复登录自动运行；`Stop.command` 停止服务并取消登录自动运行。登录时若 3000/3443 已被其他程序占用，服务直接退出、不反复重启；结束占用端口的程序后运行 `Start.command`，它会说明占用情况并重新启动。服务日志 `service.log` 超过 5MB 时自动轮换，只保留一份 `service.log.1`。
4. WPS：保存工作并重开 WPS（若弹出信任 `slidesci-mac-dev` 的提示，选择允许），从顶部 SlideSCI 分组按钮直接选择功能。
5. PowerPoint：PowerPoint 只通过受 macOS 信任的 HTTPS 加载 SlideSCI。安装时会为本机生成独立的 `localhost` 证书和私钥（有效期 825 天），保存在 `~/Library/Application Support/SlideSCI-Mac-Installer/tls/`，私钥只保存在本机。安装器会询问是否信任该证书（仅用于 `localhost` 的 SSL，不创建 CA）；同意后输入登录密码即完成。证书剩余 30 天内可运行 `RenewCertificate.command` 换新。
6. 保存工作并 `Cmd+Q` 重开 PowerPoint，从“开始 → 加载项 → 更多加载项 → 开发人员加载项”双击 `SlideSCI Mac`。如安装时未退出 PowerPoint 导致升级后功能区仍显示旧版或找不到插件，先 `Cmd+Q` 退出 PowerPoint 并删除 `~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef` 后重开即可。

ZIP 和原生辅助程序尚无 Developer ID 签名及 Apple 公证。Gatekeeper 可能阻止首次打开；确认来源后在系统设置“隐私与安全性”中允许该文件即可。浏览器下载解压后的文件都带隔离标记，未签名的原生辅助程序直接运行会被系统结束；安装器在安装副本上清除这一标记，所以只需允许一次 `Install.command`，不必逐个放行包内文件。

服务启动时不访问 PowerPoint 或 WPS 的数据，也不弹权限窗。服务与两款软件交换文件用各自沙盒允许的文件夹：PowerPoint 用 `~/Library/Application Support/Microsoft/SlideSCI-Mac`，WPS 用 `~/Library/Logs/SlideSCI-Mac`，不需要完全磁盘访问权限。首次使用要控制 PowerPoint 的功能（导出页面、复制选中对象为图片、插入格式片段与公式等）时，macOS 会询问是否允许“node”（即 SlideSCI 后台服务）控制 Microsoft PowerPoint，请选择允许；这个提示 3 分钟内无人回答会自动关闭，之后可在“系统设置 → 隐私与安全性 → 自动化”中为 node 打开 Microsoft PowerPoint。本版本已移除 AI 助手功能。

## 停止与卸载

- `Stop.command` 会通过 `launchctl bootout` 停止本机自启服务并删除登录自启配置，清理过期进程记录，不结束任何其他程序；下次登录不会自动启动，运行 `Start.command` 即恢复。
- `Uninstall.command` 先停止并删除 LaunchAgent 自启配置，再恢复安装前的宿主加载项注册。其他加载项保留。
- 卸载将程序目录移入相邻 `.uninstalled-时间` 备份目录，保留素材、设置与安装器备份，并移除安装器添加的证书信任。
- 业务数据沿用 `~/Library/Application Support/SlideSCI Mac`；安装器状态在 `~/Library/Application Support/SlideSCI-Mac-Installer`。卸载不会移除此前独立安装的开发版。

## 验证与限制

WPS 已在独立文稿中验证局部文字格式、物理比例尺、局部放大图、Markdown 原生公式及原生公式保存重开后编辑。局部颜色复制保留其余字符颜色与 22 磅字号；保存重开后比例尺长度仍为原图宽度的 20%，放大框、放大图及两条辅助线记录保留。原生公式在重开后通过 WPS 公式编辑器改写分子并保存，Office Math 结构保留。

PowerPoint 在不含旧开发加载项引用的新建文稿中完成 19 项核心实机检查，全部通过。覆盖局部文字格式、图片来源与伪彩、物理比例尺、放大图、Markdown 与原生公式、原图提取、PNG/JPEG/BMP/PDF 导出及可编辑素材保存和复用。保存重开后颜色、22 磅字号、20% 比例尺、放大图和两条辅助线记录保留；原生公式重开后经 PowerPoint 公式编辑器改写并保存，仍保留两处 Office Math。修复了 PowerPoint 粘贴重复对象时导致的 12 磅位置偏移。

隔离安装验证覆盖宿主注册与备份、生产依赖、原生辅助程序、实际内容转换、端口冲突、PID 归属、停止服务及卸载恢复。当前发行的检查结果见同一 Release 中的 `INSTALL-VERIFICATION.json`。全新 macOS 账户的 Gatekeeper、权限提示，以及安装器触发的证书信任授权窗口尚未实机验证；安装包没有 Apple 公证。

服务启动仅清理本服务在系统临时目录中残留的旧工作目录，导出完成后立即删除中间 PPTX 与 PDF。上述两个交换文件夹中的暂存文件（导出页面、粘贴图片、格式片段）在服务首次写入该文件夹时清理超过 24 小时的部分，服务持续运行时每天最多再清理一次；视频与音频保留。侧栏与服务断开后自动重连；在开启“剪贴板 / 拖入自动识别”时，服务重启后侧栏会自动恢复监听，侧栏关闭或宿主退出 60 秒后监听自动停止。

不将浏览器预览、HTTP 200 或构建成功视为宿主功能通过。字体缺失可能影响显示；源图片路径在另一台机器上未必存在；多选顺序取决于宿主返回顺序。

## 问题反馈

Mac 版问题请在 [SlideSCI-Mac Issues](https://github.com/nwh-phy/SlideSCI-Mac/issues) 反馈，附上 macOS、PowerPoint 或 WPS 版本，以及安装器状态目录中的 `service.log`（每条带时间，并记录版本与出错的操作；家目录已写作 `~`，发出前仍请检查其中的文件名）。原版 Windows 插件的作者不维护 Mac 版。

## 来源与校验

`BUILD-INFO.json` 记录版本、内置 Node 版本及归档 SHA-256、构建时间、Git 基线和源码快照指纹。该 Git 基线不代表快照中的未提交改动已公开提交。`SOURCE-MANIFEST.json` 对随包源码逐文件记录 SHA-256；`mac/` 和 `SlideSCI/latex-converter/` 包含本发行运行所需的源码及公式资源。`LICENSE` 为 AGPL-3.0；Node、生产依赖及 Phosphor 图标保留各自许可证。

同目录 `.zip.sha256` 用于核对下载文件。仅执行 `shasum -a 256 -c 文件名.zip.sha256` 不会安装软件。

开发者构建：在有 Node.js、npm 与 Xcode Command Line Tools 的 Apple Silicon Mac 上安装 `mac/` 的锁定依赖后，执行 `node mac/scripts/distribution.mjs build`。构建器校验官方 Node 归档，安装锁定生产依赖，编译原生程序并生成 ZIP 与 SHA-256。随后执行 `node mac/scripts/verify-package.mjs 安装包.zip --previous 上一版安装包.zip --report INSTALL-VERIFICATION.json`，在隔离目录中检查包内容、内置运行时、原生程序、公式转换、安装注册、过期服务记录、证书换新、卸载恢复、带隔离标记的下载副本，以及从上一版原地升级再卸载；3000/3443 端口空闲时还会检查服务启动与停止。`--strict` 把跳过的检查算作失败。仓库的 GitHub Actions（`.github/workflows/mac.yml`）在每次改动 Mac 版时运行单元测试，并在干净的 Apple Silicon 环境中构建安装包、以 `--strict` 完成上述验证（从最近一次 `mac-v*` 发布升级），验证后的 ZIP 与报告作为工作流附件保留 14 天。
