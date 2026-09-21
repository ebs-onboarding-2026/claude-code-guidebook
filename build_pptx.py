# -*- coding: utf-8 -*-
"""deck-sample.html -> deck-sample.pptx  (stdlib only, no lxml)
좌표계: HTML 1280x720 px 를 그대로 사용. 1px = 9525 EMU, 1px = 0.75pt
폰트: 맑은 고딕(Malgun Gothic) / Consolas  — Windows 기본 폰트
"""
import zipfile, os, shutil
from xml.sax.saxutils import escape

OUT = r"C:\Users\woosublee\Documents\ebs-dev\project_2\deck-sample.pptx"

EMU = 9525                      # 1px
def E(px): return int(round(px * EMU))
def SZ(px): return int(round(px * 0.75 * 100))   # px -> 1/100 pt

# ── palette ──────────────────────────────────────────────
INK   = "17150F"; INK2 = "4A463C"; MUTED = "8A8479"
PAPER = "FAF8F4"; LINE = "E2DCD1"; ACC   = "C8552F"
ACCS  = "F3E4DC"; DARK = "1A1813"; WHITE = "FFFFFF"

SANS = ('Malgun Gothic', '맑은 고딕')
MONO = ('Consolas', '맑은 고딕')

_id = [10]
def nid():
    _id[0] += 1
    return _id[0]

# ── run / paragraph ──────────────────────────────────────
def R(t, px=14, b=False, color=INK, font=SANS, spc=0, italic=False):
    return dict(t=t, sz=SZ(px), b=b, color=color, font=font, spc=spc, i=italic)

def BR():
    return dict(br=True)

def run_xml(r):
    if r.get('br'):
        return '<a:br/>'
    lat, ea = r['font']
    spc = ' spc="%d"' % int(round(r['spc'] * 100)) if r['spc'] else ''
    return (
        '<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="%d" b="%d" i="%d"%s dirty="0">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>'
        '</a:rPr><a:t>%s</a:t></a:r>'
        % (r['sz'], 1 if r['b'] else 0, 1 if r['i'] else 0, spc,
           r['color'], lat, ea, lat, escape(r['t']))
    )

def para_xml(runs, lh=None, algn='l', before=0):
    """lh = 행간(px, 고정값). PowerPoint의 배율(spcPct) 행간은 CSS line-height보다
    크게 잡히므로 항상 정확한 pt 값(spcPts)으로 지정한다."""
    pr = ''
    inner = ''
    if lh:
        inner += '<a:lnSpc><a:spcPts val="%d"/></a:lnSpc>' % int(round(lh * 0.75 * 100))
    if before:
        inner += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(before * 0.75 * 100)
    pr = '<a:pPr algn="%s">%s</a:pPr>' % (algn, inner)
    return '<a:p>' + pr + ''.join(run_xml(r) for r in runs) + '</a:p>'

# ── shapes ───────────────────────────────────────────────
def sp(x, y, w, h, geom='rect', fill=None, line=None, lw=1, dash=None,
       paras='', anchor='t', name='sp', adj=None, wrap='square'):
    g = '<a:avLst/>'
    if adj is not None:
        g = '<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>' % adj
    f = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill if fill else '<a:noFill/>'
    if line:
        d = '<a:prstDash val="%s"/>' % dash if dash else ''
        ln = ('<a:ln w="%d" cap="flat"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>'
              % (E(lw), line, d))
    else:
        ln = '<a:ln><a:noFill/></a:ln>'
    body = ('<a:bodyPr wrap="%s" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s" anchorCtr="0">'
            '<a:noAutofit/></a:bodyPr><a:lstStyle/>' % (wrap, anchor))
    if not paras:
        paras = '<a:p><a:pPr/><a:endParaRPr lang="ko-KR"/></a:p>'
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s">%s</a:prstGeom>%s%s<a:effectLst/></p:spPr>'
        '<p:txBody>%s%s</p:txBody></p:sp>'
        % (nid(), name, E(x), E(y), E(w), E(h), geom, g, f, ln, body, paras)
    )

def tb(x, y, w, h, paras, anchor='t', name='txt'):
    return sp(x, y, w, h, fill=None, line=None, paras=paras, anchor=anchor, name=name)

def rect(x, y, w, h, fill, line=None, lw=1, name='rect', geom='rect', adj=None):
    return sp(x, y, w, h, geom=geom, fill=fill, line=line, lw=lw, name=name, adj=adj)

def hline(x, y, w, color=LINE, lw=1, dash=None, name='line'):
    return sp(x, y, w, 0, geom='line', fill=None, line=color, lw=lw, dash=dash, name=name)


# ═══════════════════════════════════════════════════════════
#  SLIDE 1 — 표지
# ═══════════════════════════════════════════════════════════
def slide1():
    s = []
    s.append(rect(717, 0, 563, 720, DARK, name='rightPanel'))
    s.append(rect(717, 0, 6, 720, ACC, name='accentBar'))

    # kicker
    s.append(tb(88, 94, 560, 26, para_xml([
        R('CLAUDE CODE ', 16, True, ACC, SANS, spc=2.6),
        R('／ 입문자 가이드북', 16, True, MUTED, SANS, spc=1.7),
    ]), name='kicker'))

    # title
    s.append(tb(88, 128, 580, 182, para_xml([
        R('터미널에서', 68, True, INK, SANS, spc=-1.2), BR(),
        R('일하는 ', 68, True, INK, SANS, spc=-1.2),
        R('AI 동료', 68, True, ACC, SANS, spc=-1.2),
    ], lh=79), name='title'))

    # sub
    s.append(tb(88, 328, 480, 74, para_xml([
        R('설치부터 실전 워크플로까지,', 19, False, INK2), BR(),
        R('처음 쓰는 사람이 일주일 안에 익히는 15가지.', 19, False, INK2),
    ], lh=30), name='sub'))

    # meta
    s.append(rect(88, 558, 44, 2, INK, name='rule'))
    s.append(tb(88, 576, 420, 60, para_xml([
        R('2026 · 사내 교육자료', 14, True, MUTED, SANS, spc=0.6), BR(),
        R('Anthropic Claude Code v2', 14, True, MUTED, SANS, spc=0.6),
    ], lh=27), name='meta'))

    # terminal
    TX, TY, TW, TH = 803, 217, 392, 287
    s.append(rect(TX, TY, TW, TH, '100F0B', line='322E26', lw=1,
                  name='term', geom='roundRect', adj=3500))
    s.append(rect(TX + 1, TY + 1, TW - 2, 31, '221F19', name='termBar'))
    for i, c in enumerate(('E05A4D', 'E0B23C', '5FB868')):
        s.append(rect(TX + 13 + i * 17, TY + 11, 10, 10, c, geom='ellipse', name='dot%d' % i))

    body = ''
    body += para_xml([R('> ', 13.5, True, ACC, MONO),
                      R('이 프로젝트 구조 설명해줘', 13.5, False, 'CFC7B8', MONO)], lh=27)
    body += para_xml([R('  Read · Grep · Glob …', 13.5, False, '7E7768', MONO)], lh=27)
    body += para_xml([R('  ✓ 24개 파일 분석 완료', 13.5, False, '8FAF6E', MONO)], lh=27)
    body += para_xml([R('> ', 13.5, True, ACC, MONO),
                      R('로그인 버그 고치고 테스트해줘', 13.5, False, 'CFC7B8', MONO)], lh=27, before=10)
    body += para_xml([R('  Edit auth/login.ts', 13.5, False, '7E7768', MONO)], lh=27)
    body += para_xml([R('  ✓ 12 passed', 13.5, False, '8FAF6E', MONO)], lh=27)
    body += para_xml([R('> ', 13.5, True, ACC, MONO),
                      R('█', 13.5, False, ACC, MONO)], lh=27, before=10)
    s.append(tb(TX + 22, TY + 52, TW - 44, 210, body, name='termBody'))
    return s


# ═══════════════════════════════════════════════════════════
#  SLIDE 2 — 목차
# ═══════════════════════════════════════════════════════════
TOC = [
    ('01', '클로드 코드란?',      '챗봇과 무엇이 다른가',          False),
    ('02', '설치와 첫 실행',      'npm · 네이티브 · 로그인',       False),
    ('03', '기본 사용법',        '@ · ! · # 입력 접두사',         False),
    ('04', '핵심 단축키',        'Esc를 두려워하지 말 것',        False),
    ('05', '슬래시 명령어',      '꼭 외울 8개',                  False),
    ('06', 'CLAUDE.md',         '프로젝트 규칙 기억시키기',       False),
    ('07', '권한 모드 이해하기',  'Shift+Tab 4단계',              True),
    ('08', '플랜 모드',          '읽고 → 계획 → 승인 → 실행',     True),
    ('09', '컨텍스트 관리',      '/clear · /compact · 서브에이전트', True),
    ('10', '실전 & 확장',        'Git · 훅 · MCP · 7가지 습관',    False),
]

def slide2():
    s = []
    s.append(tb(88, 38, 640, 20, para_xml([
        R('CLAUDE CODE GUIDEBOOK ', 13, True, MUTED, SANS, spc=1.3),
        R('／', 13, True, ACC, SANS, spc=1.3),
        R(' CONTENTS', 13, True, MUTED, SANS, spc=1.3),
    ]), name='runhead'))

    s.append(tb(88, 72, 320, 26, para_xml([
        R('CONTENTS', 20, True, MUTED, SANS, spc=1.9)]), name='eyebrow'))
    # 한 줄 제목은 고정 행간을 주지 않는다 (글자가 위로 밀려 윗줄과 겹침)
    s.append(tb(88, 104, 400, 110, para_xml([
        R('목차', 80, True, INK, SANS, spc=-2.0)]), name='h1'))

    s.append(tb(692, 140, 500, 56, para_xml([
        R('기초 → 조작 → 핵심 개념 → 실전 → 확장', 14, True, MUTED), BR(),
        R('강조 항목은 ', 14, True, MUTED),
        R('입문자가 가장 많이 놓치는', 14, True, ACC),
        R(' 부분', 14, True, MUTED),
    ], lh=25, algn='r'), name='note'))

    TOP, RH = 236, 80
    COLW, GAP = 524, 56
    cols = (88, 88 + COLW + GAP)
    for i, (no, title, desc, key) in enumerate(TOC):
        cx = cols[0] if i < 5 else cols[1]
        ry = TOP + (i % 5) * RH
        if key:
            s.append(rect(cx - 12, ry, COLW + 24, 76, ACCS, name='tocBg'))
        else:
            s.append(rect(cx, ry, COLW, 1, LINE, name='tocLine'))
        s.append(tb(cx + 4, ry + 19, 34, 22, para_xml([
            R(no, 15, True, ACC, MONO)]), name='tocNo'))
        s.append(tb(cx + 46, ry + 14, COLW - 60, 28, para_xml([
            R(title, 18, True, ACC if key else INK, SANS, spc=-0.2)]), name='tocTitle'))
        s.append(tb(cx + 46, ry + 43, COLW - 60, 22, para_xml([
            R(desc, 13.5, False, MUTED)]), name='tocDesc'))

    s.append(tb(1100, 656, 124, 22, para_xml([
        R('02', 14, True, MUTED, MONO, spc=0.7)], algn='r'), name='pnum'))
    return s


# ═══════════════════════════════════════════════════════════
#  SLIDE 3 — 본문 (권한 모드)
# ═══════════════════════════════════════════════════════════
CARDS = [
    ('LEVEL 1', '기본',      'default',           'BDB6A8',
     '파일 수정과 명령 실행마다 승인을 요청합니다. 매번 눈으로 확인합니다.',
     '처음 배울 때·낯선 저장소'),
    ('LEVEL 2', '자동 수락',  'acceptEdits',       'D9A85C',
     '파일 수정은 묻지 않고 진행합니다. 명령 실행은 여전히 확인을 거칩니다.',
     '방향이 확실한 반복 작업'),
    ('LEVEL 3', '플랜',      'plan',              '5C8DB8',
     '코드를 읽기만 하고 단계별 계획을 제시합니다. 승인해야 비로소 실행합니다.',
     '큰 작업을 시작하기 전'),
    ('LEVEL 4', '전체 자동',  'bypassPermissions', ACC,
     '모든 동작을 승인 없이 실행합니다. 되돌릴 수 없는 명령까지 그대로 나갑니다.',
     '격리된 컨테이너 안에서만'),
]

def slide3():
    s = []
    s.append(tb(88, 38, 640, 20, para_xml([
        R('07', 13, True, MUTED, SANS, spc=1.3),
        R(' ／ ', 13, True, ACC, SANS, spc=1.3),
        R('권한 모드 이해하기', 13, True, MUTED, SANS, spc=1.3),
    ]), name='runhead'))

    s.append(rect(88, 101, 26, 3, ACC, name='eyeBar'))
    s.append(tb(124, 92, 320, 22, para_xml([
        R('SECTION 07', 15, True, ACC, SANS, spc=1.35)]), name='eyebrow'))

    s.append(tb(88, 124, 1104, 64, para_xml([
        R('얼마나 맡길지, ', 46, True, INK, SANS, spc=-0.7),
        R('Shift + Tab', 46, True, ACC, SANS, spc=-0.7),
        R(' 한 번으로 정한다', 46, True, INK, SANS, spc=-0.7),
    ], lh=57), name='title'))

    s.append(tb(88, 198, 950, 72, para_xml([
        R('클로드가 파일을 고치고 명령을 실행하기 전에 ', 20, False, INK2),
        R('어디까지 물어볼지', 20, True, INK),
        R('를 4단계로 조절합니다. 낯선 저장소일수록 왼쪽, 방향이 확실할수록 오른쪽.', 20, False, INK2),
    ], lh=32), name='lead'))

    CY, CH, CW, CG = 290, 272, 264, 16
    for i, (lv, ko, en, bar, body, when) in enumerate(CARDS):
        cx = 88 + i * (CW + CG)
        s.append(rect(cx, CY, CW, CH, WHITE, line=LINE, lw=1, name='card'))
        s.append(rect(cx, CY, CW, 4, bar, name='cardBar'))
        ix, iw = cx + 22, CW - 44
        s.append(tb(ix, CY + 24, iw, 18, para_xml([
            R(lv, 12, True, MUTED, MONO, spc=0.9)]), name='lv'))
        s.append(tb(ix, CY + 48, iw, 30, para_xml([
            R(ko, 20, True, INK, SANS, spc=-0.3)]), name='ko'))
        s.append(tb(ix, CY + 80, iw, 20, para_xml([
            R(en, 12.5, True, ACC, MONO)]), name='en'))
        # 본문 3줄(25px) = 75px → 점선까지 11px 여유
        s.append(tb(ix, CY + 110, iw, 78, para_xml([
            R(body, 14.5, False, INK2)], lh=25), name='body'))
        s.append(hline(ix, CY + 196, iw, LINE, 1, dash='dash', name='dash'))
        s.append(tb(ix, CY + 206, iw, 16, para_xml([
            R('WHEN', 11, True, MUTED, SANS, spc=1.1)]), name='whenLabel'))
        s.append(tb(ix, CY + 224, iw, 24, para_xml([
            R(when, 13, True, MUTED)]), name='when'))

    # callout
    s.append(rect(88, 586, 1104, 70, DARK, name='callout'))
    s.append(sp(114, 606, 30, 30, geom='ellipse', fill=ACC, name='mark',
                paras=para_xml([R('!', 17, True, WHITE)], algn='ctr'), anchor='ctr'))
    s.append(tb(162, 586, 1006, 70, para_xml([
        R('귀찮다고 모드를 올리지 말고 ', 15.5, True, WHITE),
        R('/permissions', 14, True, 'E8A98E', MONO),
        R('를 쓰세요. ', 15.5, True, WHITE),
        R('Bash(npm test:*)', 14, False, 'E8A98E', MONO),
        R('처럼 안전한 명령만 허용 목록에 넣는 편이 훨씬 낫습니다.', 15.5, False, 'EFE9DF'),
    ], lh=25), anchor='ctr', name='calloutTx'))

    s.append(tb(1100, 656, 124, 22, para_xml([
        R('14', 14, True, MUTED, MONO, spc=0.7)], algn='r'), name='pnum'))
    return s


# ═══════════════════════════════════════════════════════════
#  OOXML 패키징
# ═══════════════════════════════════════════════════════════
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')

def slide_xml(shapes, bg=PAPER):
    return (HDR + '<p:sld %s><p:cSld>' % NS +
            '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:effectLst/></p:bgPr></p:bg>' % bg +
            '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
            ''.join(shapes) +
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')

CT = HDR + '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS = HDR + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

PRES = HDR + '''<p:presentation %s saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
<p:defaultTextStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr></p:defaultTextStyle>
</p:presentation>''' % NS

PRES_RELS = HDR + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>'''

EMPTY_TREE = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
              '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>')

TXSTYLE = ''.join(
    '<p:%sStyle>' % k + ''.join(
        '<a:lvl%dpPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl%dpPr>' % (i, i)
        for i in range(1, 10)) + '</p:%sStyle>' % k
    for k in ('title', 'body', 'other'))

MASTER = HDR + ('<p:sldMaster %s><p:cSld>'
    '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    '%s</p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
    'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" '
    'folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles>%s</p:txStyles></p:sldMaster>' % (NS, PAPER, EMPTY_TREE, TXSTYLE))

MASTER_RELS = HDR + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''

LAYOUT = HDR + ('<p:sldLayout %s type="blank" preserve="1"><p:cSld name="빈 화면">%s</p:cSld>'
                '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>' % (NS, EMPTY_TREE))

LAYOUT_RELS = HDR + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

SLIDE_RELS = HDR + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

def _clr(tag, v):
    return '<a:%s><a:srgbClr val="%s"/></a:%s>' % (tag, v, tag)

THEME = HDR + ('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ClaudeCodeDeck">'
    '<a:themeElements><a:clrScheme name="ClaudeCode">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    + _clr('dk2', INK) + _clr('lt2', PAPER)
    + _clr('accent1', ACC) + _clr('accent2', 'D9A85C') + _clr('accent3', '5C8DB8')
    + _clr('accent4', 'BDB6A8') + _clr('accent5', '8FAF6E') + _clr('accent6', '7E7768')
    + _clr('hlink', ACC) + _clr('folHlink', MUTED) +
    '</a:clrScheme>'
    '<a:fontScheme name="Malgun">'
    '<a:majorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme>'
    '<a:fmtScheme name="Simple">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="28575" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme></a:themeElements></a:theme>')

CORE = HDR + ('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>클로드 코드 입문자 가이드북 — 샘플</dc:title>'
    '<dc:creator>Claude Code</dc:creator><cp:lastModifiedBy>Claude Code</cp:lastModifiedBy>'
    '</cp:coreProperties>')

APP = HDR + ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Application>Microsoft Office PowerPoint</Application><Slides>3</Slides>'
    '<PresentationFormat>와이드스크린</PresentationFormat></Properties>')


def build():
    if os.path.exists(OUT):
        os.remove(OUT)
    z = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', ROOT_RELS)
    z.writestr('docProps/core.xml', CORE)
    z.writestr('docProps/app.xml', APP)
    z.writestr('ppt/presentation.xml', PRES)
    z.writestr('ppt/_rels/presentation.xml.rels', PRES_RELS)
    z.writestr('ppt/theme/theme1.xml', THEME)
    z.writestr('ppt/slideMasters/slideMaster1.xml', MASTER)
    z.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', MASTER_RELS)
    z.writestr('ppt/slideLayouts/slideLayout1.xml', LAYOUT)
    z.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', LAYOUT_RELS)
    for i, (shapes, bg) in enumerate(((slide1(), PAPER), (slide2(), PAPER), (slide3(), PAPER)), 1):
        z.writestr('ppt/slides/slide%d.xml' % i, slide_xml(shapes, bg))
        z.writestr('ppt/slides/_rels/slide%d.xml.rels' % i, SLIDE_RELS)
    z.close()
    print('written:', OUT, os.path.getsize(OUT), 'bytes')

if __name__ == '__main__':
    build()
