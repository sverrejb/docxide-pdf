"""Generate synthetic U+2011 fallback test fonts. Requires fontTools 4.67.0.
No system font data is used; outputs are deterministic.
"""
from pathlib import Path
from fontTools.fontBuilder import FontBuilder
from fontTools.ttLib import newTable
from fontTools.ttLib.tables._k_e_r_n import KernTable_format_0
from fontTools.pens.ttGlyphPen import TTGlyphPen
import argparse


def rect(pen,x,y,w,h):
    pen.moveTo((x,y));pen.lineTo((x+w,y));pen.lineTo((x+w,y+h));pen.lineTo((x,y+h));pen.closePath()


def build(path,kind):
    fb=FontBuilder(1000,isTTF=True)
    names=['.notdef','space','one','two','hyphen','unicodeHyphen','nonbreakingHyphen']
    fb.setupGlyphOrder(names)
    glyphs={}
    for name in names:
        p=TTGlyphPen(None)
        if name=='.notdef':rect(p,50,0,450,650)
        elif name=='one':rect(p,200,0,80,650)
        elif name=='two':
            for box in [(60,570,380,80),(360,325,80,300),(60,285,380,80),(60,0,80,300),(60,0,380,80)]:rect(p,*box)
        elif name=='hyphen':rect(p,40,280,300,70)
        elif name=='unicodeHyphen':rect(p,40,280,400,70)
        elif name=='nonbreakingHyphen':rect(p,40,280,600,70)
        glyphs[name]=p.glyph()
    cmap={32:'space',49:'one',50:'two'}
    if kind!='none':cmap[45]='hyphen'
    if kind in ['unicode','native']:cmap[0x2010]='unicodeHyphen'
    if kind=='native':cmap[0x2011]='nonbreakingHyphen'
    fb.setupCharacterMap(cmap);fb.setupGlyf(glyphs)
    advances={n:600 for n in names};advances.update(space=250,hyphen=400,unicodeHyphen=500,nonbreakingHyphen=700)
    fb.setupHorizontalMetrics({n:(advances[n],0) for n in names})
    fb.setupHorizontalHeader(ascent=800,descent=-200)
    family='NB Hyphen Fixture '+kind
    fb.setupNameTable({'familyName':family,'styleName':'Regular','uniqueFontIdentifier':family+' 1.0','fullName':family,'psName':family.replace(' ',''),'version':'Version 1.0'})
    fb.setupOS2(sTypoAscender=800,sTypoDescender=-200,usWinAscent=800,usWinDescent=200,fsType=0)
    kern=newTable('kern');kern.version=0
    pairs=KernTable_format_0();pairs.version=0;pairs.coverage=1
    pairs.kernTable={('hyphen','one'):-50,('unicodeHyphen','one'):-75,('nonbreakingHyphen','one'):-100}
    kern.kernTables=[pairs];fb.font['kern']=kern
    fb.setupPost();fb.setupMaxp();fb.font['head'].created=fb.font['head'].modified=3864412800;fb.font.recalcTimestamp=False
    fb.save(path)


def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();a.out.mkdir(parents=True,exist_ok=True)
    for kind in ['ascii','unicode','native','none']:build(a.out/('nb-hyphen-'+kind+'.ttf'),kind)


if __name__=='__main__':main()
