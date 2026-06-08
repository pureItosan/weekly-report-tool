# -*- coding: utf-8 -*-
"""Inline styles.css, vendor JS and app.js into ONE self-contained HTML file
that works by double-click (file://), no server needed."""
import os, re
base=os.path.dirname(os.path.abspath(__file__))
def read(p):
    with open(os.path.join(base,p),encoding='utf-8') as f: return f.read()

html=read('index.html')
# strip cache-busting query strings (?v=...) so the inliner below matches the tags and the
# bundle ends up fully self-contained (no external refs remain)
html=re.sub(r'(href="styles\.css|src="app\.js)\?v=[^"]*"', r'\1"', html)
# the online page loads app.js via a dynamic loader (cache-bust); for the single file, replace
# that loader with a plain placeholder so the inliner below can drop the app code in its place.
html=re.sub(r'<script id="app-loader">.*?</script>', '<script src="app.js"></script>', html, flags=re.DOTALL)
css=read('styles.css')
jszip=read('vendor/jszip.min.js')
xlsx=read('vendor/xlsx.full.min.js')
pptx=read('vendor/pptxgen.bundle.js')
app=read('app.js')

def safe(js):  # avoid premature </script> termination
    return js.replace('</script','<\\/script')

# replace <link rel=stylesheet ...> with inline style
html=re.sub(r'<link rel="stylesheet" href="styles.css">',
            '<style>\n'+css+'\n</style>', html)
# replace the three external scripts with inline versions
html=html.replace('<script src="vendor/jszip.min.js"></script>',
                  '<script>'+safe(jszip)+'</script>')
html=html.replace('<script src="vendor/xlsx.full.min.js"></script>',
                  '<script>'+safe(xlsx)+'</script>')
html=html.replace('<script src="vendor/pptxgen.bundle.js"></script>',
                  '<script>'+safe(pptx)+'</script>')
html=html.replace('<script src="app.js"></script>',
                  '<script>'+safe(app)+'</script>')

out=os.path.join(base,'週報整合工具.html')
with open(out,'w',encoding='utf-8') as f: f.write(html)
print('wrote', out, round(os.path.getsize(out)/1024), 'KB')
# sanity: ensure no leftover external refs
for tag in ['href="styles.css"','src="vendor','src="app.js"']:
    assert tag not in html, 'leftover '+tag
print('no external references remain — fully self-contained')
