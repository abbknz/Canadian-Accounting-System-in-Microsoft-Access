"""Puts every picture of the getting-started guide into docs/images and lists them, with their descriptions, in README.md.
Usage: make_readme_screens.py <project folder>   (run tools/make_guide.py first)"""
import html, os, re, shutil, sys
from PIL import Image

proj = sys.argv[1]
guide = os.path.join(proj, 'user-guide')
out = os.path.join(proj, 'docs', 'images')
if os.path.isdir(out):
    shutil.rmtree(out)
os.makedirs(out)

page = open(os.path.join(guide, 'Kanzagh_Accounting_Getting_Started.html'), encoding='utf-8').read()
page = re.sub(r'base64,[A-Za-z0-9+/=]+', '', page)
tokens = re.findall(r'<h2 id="p\d+">(.*?)</h2>|<h3><span>Step (\d+)\.</span> (.*?)</h3><p>(.*?)</p><img', page, re.S)
plain = lambda s: html.unescape(re.sub(r'<[^>]+>', '', s)).strip()
lines = ['## Screens', '',
         'Every screen below was captured in the program while a sample company, Maple Ridge Outfitters Ltd., was set up and its first three months were entered. The company and the figures are sample data.', '']
count = 0
for part, num, title, text in tokens:
    if part:
        lines += ['### ' + plain(part), '']
        continue
    title, text = plain(title), plain(text)
    name = '%02d-%s.png' % (int(num), re.sub(r'[^a-z0-9]+', '-', title.lower()).strip('-'))
    im = Image.open(os.path.join(guide, 'images', '%02d.png' % int(num)))
    im.crop((10, 0, im.width, im.height)).save(os.path.join(out, name))        # the window's dark left edge is cut off
    lines += ['**%d. %s**' % (int(num), title), '', text, '', '![%s](docs/images/%s)' % (title, name), '']
    count += 1

path = os.path.join(proj, 'README.md')
readme = open(path, encoding='utf-8').read()
readme = re.sub(r'\n## Screens\n.*?(?=\n## |\Z)', '\n', readme, flags=re.S).rstrip() + '\n'
marker = '\n## What is in this repository'
assert readme.count(marker) == 1
readme = readme.replace(marker, '\n' + '\n'.join(lines).rstrip() + '\n' + marker)
open(path, 'w', encoding='utf-8', newline='\n').write(readme)
print('pictures:', count, '| README.md', len(readme), 'characters')
