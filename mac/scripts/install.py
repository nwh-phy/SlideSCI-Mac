#!/usr/bin/env python3
"""Register only the SlideSCI development add-in; preserve existing registrations."""
import argparse
from datetime import datetime
from pathlib import Path
import shutil
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument('host', choices=['wps', 'powerpoint'])
args = parser.parse_args()
source = Path(__file__).resolve().parents[1]
container = Path.home() / 'Library/Containers'
stamp = datetime.now().strftime('%Y%m%d-%H%M%S-%f')

def backup(path):
    if path.exists():
        target = path.with_name(path.name + '.backup-' + stamp)
        if path.is_dir():
            shutil.copytree(path, target)
        else:
            shutil.copy2(path, target)
        print('备份:', target)

if args.host == 'powerpoint':
    base = container / 'com.microsoft.Powerpoint/Data/Documents'
    if not base.is_dir():
        parser.error('未找到 PowerPoint 容器，请先安装并打开 PowerPoint。')
    target = base / 'wef/SlideSCI-Mac.xml'
    target.parent.mkdir(exist_ok=True)
    backup(target)
    shutil.copy2(source / 'office/manifest.xml', target)
    print('已注册:', target)
    print('需启动 HTTPS 服务并信任开发证书，再从 PowerPoint 的加载项菜单打开。')
else:
    base = container / 'com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons'
    if not base.is_dir():
        parser.error('未找到 WPS 加载项目录，请先安装并打开 WPS。')
    publish = base / 'publish.xml'
    tree = ET.parse(publish) if publish.exists() else ET.ElementTree(ET.Element('jsplugins'))
    root = tree.getroot()
    if root.tag != 'jsplugins':
        parser.error('publish.xml 根节点未知，停止写入。')
    name = 'slidesci-mac-dev'
    entries = [entry for entry in root if entry.get('name') == name]
    if len(entries) > 1:
        parser.error('发现重复 SlideSCI 注册，停止写入。')
    backup(publish)
    entry = entries[0] if entries else ET.SubElement(root, 'jspluginonline')
    entry.tag = 'jspluginonline'
    entry.attrib.clear()
    entry.attrib.update(name=name, type='wpp', url='http://127.0.0.1:3000/wps/',
                        enable='enable_dev', install='null')
    temporary = publish.with_name('publish.xml.slidesci-tmp')
    tree.write(temporary, encoding='utf-8', xml_declaration=True)
    temporary.replace(publish)
    print('已注册:', name, '→', entry.get('url'))
    print('实际启用取决于当前 WPS Mac 版本；不会修改 WPS 安全开关或安装包。')
