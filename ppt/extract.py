# Extract slide data from deck.html into ppt/slides.json and image sizes into ppt/sizes.json
import re, json, html, os
from PIL import Image
s = open('deck.html', encoding='utf-8').read()
def txt(x): return html.unescape(re.sub(r'<[^>]+>', '', x)).strip()
out = []
for m in re.finditer(r'<section class="slide[^"]*" data-title="([^"]*)"[^>]*>(.*?)</section>', s, re.S):
    title, body = m.group(1), m.group(2)
    eb = re.search(r'<div class="eyebrow"[^>]*>(.*?)</div>', body, re.S)
    h2 = re.search(r'<h[12][^>]*>(.*?)</h[12]>', body, re.S)
    lede = re.search(r'<p class="lede">(.*?)</p>', body, re.S)
    steps = [dict(img=a, cap=html.unescape(b), title=txt(c), desc=txt(d)) for a, b, c, d in
             re.findall(r'<li data-img="([^"]*)" data-cap="([^"]*)"><span class="n">\d+</span><div><b>(.*?)</b><p>(.*?)</p></div></li>', body, re.S)]
    prompts = [dict(lab=txt(re.sub(r'<button.*?</button>', '', a)), text=txt(b)) for a, b in
               re.findall(r'<div class="prompt"><div class="lab">(.*?)</div>(.*?)</div>', body, re.S)]
    out.append(dict(title=title, eyebrow=txt(eb.group(1)) if eb else '', h2=txt(h2.group(1)) if h2 else '',
                    lede=txt(lede.group(1)) if lede else '', steps=steps, prompts=prompts))
json.dump(out, open('ppt/slides.json', 'w'), ensure_ascii=False, indent=1)
sizes = {}
for p in sorted(set(re.findall(r'(img/[\w\-.]+\.(?:jpg|png))', s))):
    if os.path.exists(p): sizes[p] = list(Image.open(p).size)
json.dump(sizes, open('ppt/sizes.json', 'w'), indent=0)
print(len(out), 'slides,', len(sizes), 'images')
