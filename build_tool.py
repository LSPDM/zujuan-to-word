# -*- coding: utf-8 -*-
"""把 src/ 下的模板与脚本打包成单个离线 HTML 工具。

用法：
    python build_tool.py                # 输出到 dist/（index.html 与中文名各一份）
    python build_tool.py 输出路径.html   # 输出到指定文件
"""
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, 'src')
DIST = os.path.join(ROOT, 'dist')
HTML_NAME = '组卷转Word工具.html'


def build(out_path):
    tpl = open(os.path.join(SRC, 'index.tpl.html'), encoding='utf-8').read()

    parts = []
    for name in ('core.js', 'docx.js', 'ui.js'):
        code = open(os.path.join(SRC, name), encoding='utf-8').read()
        parts.append('<script>\n/* ===== %s ===== */\n%s\n</script>' % (name, code))

    html = tpl.replace('<!--SCRIPTS-->', '\n'.join(parts))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    # 统一用 LF 换行，保证本地与 CI 构建结果逐字节一致
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(html)
    return os.path.getsize(out_path)


def main():
    if len(sys.argv) > 1:
        targets = [os.path.abspath(sys.argv[1])]
    else:
        targets = [os.path.join(DIST, HTML_NAME), os.path.join(DIST, 'index.html')]
    for t in targets:
        size = build(t)
        print('built %s  (%.1f KB)' % (t, size / 1024))


if __name__ == '__main__':
    main()
