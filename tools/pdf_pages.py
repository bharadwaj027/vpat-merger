import sys, fitz
doc = fitz.open(sys.argv[1])
print('pages', doc.page_count)
for n in [int(x) for x in sys.argv[3].split(',')]:
    pix = doc[n-1].get_pixmap(dpi=70)
    pix.save(f"{sys.argv[2]}/page-{n}.png")
