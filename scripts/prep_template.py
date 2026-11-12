"""Convert the .potx to a python-pptx-openable .pptx and list layout placeholders."""
import zipfile, sys
from pptx import Presentation
from pptx.util import Emu

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == '[Content_Types].xml':
        data = data.replace(b'presentationml.template.main+xml', b'presentationml.presentation.main+xml')
    zout.writestr(item, data)
zout.close()

p = Presentation(dst)
print('slide size', Emu(p.slide_width).inches, Emu(p.slide_height).inches, 'slides', len(p.slides))
for i, l in enumerate(p.slide_layouts, 1):
    print(f'--- {i} {l.name}')
    for ph in l.placeholders:
        f = ph.placeholder_format
        print(f'   idx={f.idx} type={f.type} name={ph.name!r} '
              f'pos=({Emu(ph.left).inches:.2f},{Emu(ph.top).inches:.2f}) size=({Emu(ph.width).inches:.2f}x{Emu(ph.height).inches:.2f})')
