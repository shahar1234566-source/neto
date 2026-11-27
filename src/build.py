"""Builds app-index.html (uploaded to the Drive folder 'נטו הזמנות') from client.html + style.css.part.
The main script is shipped base64-encoded because Apps Script HtmlService breaks some inline JS on iPhone."""
import base64, sys, os
here = os.path.dirname(os.path.abspath(__file__))
c = open(os.path.join(here, 'client.html'), encoding='utf-8').read()
start = c.index('<script>\nconst BOOT'); end = c.index('</script>', start)
js = c[start + len('<script>\n'):end].replace('const BOOT=__BOOT__;', 'const BOOT=window.__BOOT;')
b64 = base64.b64encode(js.encode()).decode()
loader = ('<script>window.__BOOT=__BOOT__;</script>\n<script>\n(function(){try{var b=atob("%s");var u=new Uint8Array(b.length);'
          'for(var i=0;i<b.length;i++)u[i]=b.charCodeAt(i);var s=document.createElement("script");'
          's.text=new TextDecoder("utf-8").decode(u);document.body.appendChild(s);}catch(e){window.__errs.push("loader: "+e.message);showDiag();}})();\n</script>') % b64
out = (c[:start] + loader + c[end + 9:]).replace('__STYLE__', open(os.path.join(here, 'style.css.part'), encoding='utf-8').read())
dst = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, 'app-index.html')
open(dst, 'w', encoding='utf-8').write(out)
print('wrote', dst, len(out))
