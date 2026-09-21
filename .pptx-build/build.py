from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import json

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'deck-sample.pptx'
NS='http://schemas.openxmlformats.org'
A=f'{NS}/drawingml/2006/main'
P=f'{NS}/presentationml/2006/main'
R=f'{NS}/officeDocument/2006/relationships'
BLUE='2457FF'; INK='141821'; MUTED='69717E'
def emu(v): return round(v*9525)
slides=[]; blocks=[]
def text(x,y,w,h,s,size=20,color=INK,bold=False,font='Malgun Gothic',leading=120):
    i=len(blocks)+2
    paras=[]
    for line in s.split('\n'):
        paras.append(f'<a:p><a:pPr><a:lnSpc><a:spcPct val="{leading*1000}"/></a:lnSpc><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr><a:r><a:rPr lang="ko-KR" sz="{round(size*75)}" b="{int(bold)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="{font}"/></a:rPr><a:t>{escape(line)}</a:t></a:r><a:endParaRPr lang="ko-KR" sz="{round(size*75)}"/></a:p>')
    blocks.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Text {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def rect(x,y,w,h,color):
    i=len(blocks)+2
    blocks.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Panel {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>')
def brand():
    text(68,45,190,28,'claude code guide',18,BLUE,True)
def footer(n,s):
    text(68,666,1080,25,s,13,'7B8391');text(1174,664,40,27,f'{n:02}',15,BLUE,True)
def finish(notes):
    global blocks
    slides.append((''.join(blocks),notes));blocks=[]

brand()
text(68,151,660,32,'처음 시작하는 AI 코딩',18,BLUE,True)
text(68,204,720,90,'클로드 코드',64,INK,True)
text(68,290,720,94,'입문자 가이드북',64,BLUE,True)
text(68,421,650,82,'설치부터 첫 번째 결과물까지\n직접 따라 하며 배우는 Claude Code',23,MUTED,leading=150)
text(68,563,620,30,'BEGINNER’S GUIDE   2026',15,'7B8391')
rect(845,174,367,3,BLUE)
text(845,199,367,218,'Hello,\nClaude.',92,BLUE,True,font='Segoe UI',leading=110)
text(845,444,365,90,'나의 첫 AI\n코딩 프로젝트',29,INK,True,leading=135)
text(845,552,365,60,'한국어로 요청하고\n내 프로젝트에서 결과를 확인합니다.',16,MUTED,leading=145)
footer(1,'설치 · 실습 · 작업 습관')
finish('원문: 클로드코드-입문자-가이드북.md\n디자인 참고: https://bysuppt.com/work/jgemvi?fromPage=1&fromScroll=1500\nHTML 샘플의 흰색과 파란색 구성을 편집 가능한 개체로 재구성했습니다.')

brand();text(927,49,290,24,'01  클로드 코드 이해하기',14,'8B92A0')
text(68,111,1120,57,'내 프로젝트에서 함께 작업하는',41,INK,True)
text(68,167,1120,57,'AI 코딩 도구',41,BLUE,True)
text(68,237,1120,35,'Claude Code는 파일을 읽고, 코드를 수정하고, 필요한 명령을 실행합니다.',20,MUTED)
for j,(title,desc) in enumerate([
 ('프로젝트 파악','폴더 구조와 코드의 역할을 읽고\n초보자도 이해할 수 있게 설명합니다.'),
 ('코드 수정','한국어로 설명한 요구사항을 바탕으로\n기능을 추가하거나 오류를 고칩니다.'),
 ('결과 확인','실행과 테스트를 요청하고,\n완성한 기능은 직접 확인합니다.')]):
    y=312+j*108
    text(68,y,40,32,f'{j+1:02}',19,BLUE,True)
    text(118,y-2,450,37,title,26,INK,True)
    text(118,y+40,475,57,desc,18,MUTED,leading=140)
rect(661,304,551,321,'F4F6FB')
text(694,329,480,30,'처음 입력해 볼 요청',15,BLUE,True)
text(694,373,480,125,'“이 프로젝트의 목적과\n폴더 구조를 초보자에게\n설명해줘.”',27,INK,True,leading=140)
rect(694,506,480,1,'DCE2EE')
text(694,524,485,53,'첫 요청에는 이렇게 덧붙여 보세요.\n“수정은 하지 말고 실행 방법부터 알려줘.”',17,MUTED,leading=140)
text(694,590,485,28,'작게 요청하고, 결과를 확인하기',19,BLUE,True)
footer(2,'출처: Claude Code 공식 Quickstart · 원문 가이드 기준 2026.09.14')
finish('제품 설명 출처: https://code.claude.com/docs/en/quickstart\n원문 가이드 기준일: 2026-09-14. 예시 요청문은 학습을 위해 작성한 문구입니다.')

brand();text(927,49,290,24,'02  설치하고 로그인하기',14,'8B92A0')
text(68,111,1110,64,'Windows에서 설치와 첫 실행',42,INK,True)
text(68,184,1120,34,'시작 메뉴에서 PowerShell을 열고, 아래 명령을 순서대로 입력합니다.',20,MUTED)
def command(y,n,label,code,note=None):
    text(68,y,42,29,n,18,BLUE,True);text(108,y,700,29,label,18,INK,True)
    lines=len(code.split('\n'));h=30*lines+22
    rect(68,y+36,723,h,'F2F5FC');rect(68,y+36,3,h,BLUE)
    text(87,y+46,690,h-10,code,20,'24365C',font='Consolas',leading=145)
    if note:text(68,y+43+h,723,27,note,15,MUTED)
command(249,'01','공식 설치 스크립트 실행','irm https://claude.ai/install.ps1 | iex','공식 서버에서 설치 스크립트를 내려받아 실행합니다.')
command(380,'02','새 PowerShell 창에서 설치 확인','claude --version')
command(484,'03','실습 폴더를 만들고 시작','mkdir claude-first-project\ncd claude-first-project\nclaude')
rect(832,249,1,374,'E5E8EF')
text(866,249,340,36,'시작 전 확인',24,INK,True)
text(866,300,340,28,'계정과 인터넷 연결',18,INK,True)
text(866,336,340,82,'사용 권한이 있는 계정이 필요합니다.\n구독 이용과 Console API 과금은\n구분해 확인하세요.',17,MUTED,leading=140)
rect(866,429,340,1,'E5E8EF')
text(866,447,340,28,'첫 실행 후 로그인',18,INK,True)
text(866,480,340,54,'화면 안내에 따라 로그인하고,\n현재 작업 폴더를 확인합니다.',17,MUTED,leading=140)
text(866,553,340,29,'입력 위치 구분',18,BLUE,True)
text(866,586,340,61,'명령은 PowerShell에, 한국어 요청은\nClaude Code 입력창에 씁니다.',17,MUTED,leading=140)
footer(3,'기존 폴더와 이름이 같으면 다른 이름을 사용하세요. · 공식 문서: code.claude.com/docs/en/setup')
finish('설치 및 계정 출처: https://code.claude.com/docs/en/setup\n설치 명령 출처: https://code.claude.com/docs/en/quickstart\n원문 기준일: 2026-09-14. 버전 및 조직 설정에 따라 화면이 다를 수 있습니다. 설치 명령은 공식 스크립트를 다운로드하여 실행합니다.')

def rels(items):
    return f'<Relationships xmlns="{NS}/package/2006/relationships">'+''.join(f'<Relationship Id="{i}" Type="{R}/{t}" Target="{target}"/>' for i,t,target in items)+'</Relationships>'
grp='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
files={}
files['_rels/.rels']=rels([('rId1','officeDocument','ppt/presentation.xml')])
files['ppt/presentation.xml']=f'<p:presentation xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId4"/></p:sldMasterIdLst><p:sldIdLst>'+''.join(f'<p:sldId id="{255+i}" r:id="rId{i}"/>' for i in range(1,4))+f'</p:sldIdLst><p:sldSz cx="{emu(1280)}" cy="{emu(720)}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
files['ppt/_rels/presentation.xml.rels']=rels([(f'rId{i}','slide',f'slides/slide{i}.xml') for i in range(1,4)]+[('rId4','slideMaster','slideMasters/slideMaster1.xml')])
files['ppt/slideMasters/slideMaster1.xml']=f'<p:sldMaster xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree>{grp}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'
files['ppt/slideMasters/_rels/slideMaster1.xml.rels']=rels([('rId1','slideLayout','../slideLayouts/slideLayout1.xml'),('rId2','theme','../theme/theme1.xml')])
files['ppt/slideLayouts/slideLayout1.xml']=f'<p:sldLayout xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>{grp}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
files['ppt/slideLayouts/_rels/slideLayout1.xml.rels']=rels([('rId1','slideMaster','../slideMasters/slideMaster1.xml')])
colors={'dk1':INK,'lt1':'FFFFFF','dk2':MUTED,'lt2':'F4F6FB','accent1':BLUE,'accent2':'2457FF','accent3':'69717E','accent4':'DCE2EE','accent5':'24365C','accent6':'EFF3FF','hlink':BLUE,'folHlink':'69717E'}
files['ppt/theme/theme1.xml']=f'<a:theme xmlns:a="{A}" name="System Blue"><a:themeElements><a:clrScheme name="Blue">'+''.join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k,v in colors.items())+'</a:clrScheme><a:fontScheme name="System"><a:majorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Segoe UI"/></a:majorFont><a:minorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Segoe UI"/></a:minorFont></a:fontScheme><a:fmtScheme name="Simple"><a:fillStyleLst>'+''.join('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' for _ in range(3))+'</a:fillStyleLst><a:lnStyleLst>'+''.join('<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' for _ in range(3))+'</a:lnStyleLst><a:effectStyleLst>'+''.join('<a:effectStyle><a:effectLst/></a:effectStyle>' for _ in range(3))+'</a:effectStyleLst><a:bgFillStyleLst>'+''.join('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' for _ in range(3))+'</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
types=[('/ppt/presentation.xml','presentation.main'),('/ppt/slideMasters/slideMaster1.xml','slideMaster'),('/ppt/slideLayouts/slideLayout1.xml','slideLayout')]
for i,(content,notes) in enumerate(slides,1):
    files[f'ppt/slides/slide{i}.xml']=f'<p:sld xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>{grp}{content}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    files[f'ppt/slides/_rels/slide{i}.xml.rels']=rels([('rId1','slideLayout','../slideLayouts/slideLayout1.xml'),('rId2','notesSlide',f'../notesSlides/notesSlide{i}.xml')])
    files[f'ppt/notesSlides/notesSlide{i}.xml']=f'<p:notes xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree>{grp}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{escape(notes)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>'
    files[f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels']=rels([('rId1','slide',f'../slides/slide{i}.xml')])
    types.extend([(f'/ppt/slides/slide{i}.xml','slide'),(f'/ppt/notesSlides/notesSlide{i}.xml','notesSlide')])
files['[Content_Types].xml']=f'<Types xmlns="{NS}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'+''.join(f'<Override PartName="{part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.{kind}+xml"/>' for part,kind in types)+'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
    for name,content in files.items():z.writestr(name,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+content)
print(OUT)
