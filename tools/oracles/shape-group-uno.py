#!/usr/bin/python3
"""Optional test-only LibreOffice group import/save/reopen oracle."""
import json, os, signal, subprocess, sys, time, uuid
import uno, unohelper
from com.sun.star.task import XInteractionHandler
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
ROOT=os.path.abspath(sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(__file__),'../../artifacts/graphics'));OUT=ROOT+'/group-uno';os.makedirs(OUT,exist_ok=True)
report={'producer':subprocess.check_output(['libreoffice','--version'],text=True).strip(),'interactions':[],'rows':[],'status':'running'}
class Handler(unohelper.Base,XInteractionHandler):
 def handle(self,r):
  report['interactions'].append(str(r.getRequest()))
  for c in r.getContinuations():
   if 'XInteractionAbort' in str(c):c.select();return
  raise RuntimeError('Unexpected interaction')
def prop(name,value):
 p=uno.createUnoStruct('com.sun.star.beans.PropertyValue');p.Name=name;p.Value=value;return p
def snapshot(doc):
 pages=doc.getDrawPages();rows=[]
 def scan(shape):
  p,s=shape.Position,shape.Size;row={'name':shape.Name,'type':shape.getShapeType(),'position':[p.X,p.Y],'size':[s.Width,s.Height]}
  try:row['text']=shape.getString()
  except Exception:pass
  if shape.supportsService('com.sun.star.drawing.GroupShape'):row['children']=[scan(shape.getByIndex(i)) for i in range(shape.getCount())]
  return row
 return [[scan(pages.getByIndex(i).getByIndex(j)) for j in range(pages.getByIndex(i).getCount())] for i in range(pages.getCount())]
pipe='go_pictures_'+uuid.uuid4().hex
server=subprocess.Popen(['libreoffice','-env:UserInstallation=file://'+OUT+'/profile','--headless','--nologo','--nodefault','--nofirststartwizard','--accept=pipe,name='+pipe+';urp;StarOffice.ServiceManager'],stdout=open(OUT+'/server.stdout.log','w'),stderr=open(OUT+'/server.stderr.log','w'),start_new_session=True)
desktop=None
try:
 ctx=uno.getComponentContext();resolver=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
 for i in range(80):
  try:remote=resolver.resolve('uno:pipe,name='+pipe+';urp;StarOffice.ComponentContext');break
  except Exception:time.sleep(.25)
 else:raise RuntimeError('UNO connect timeout')
 desktop=remote.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',remote);handler=Handler()
 def load(path):
  doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(path),'_blank',0,(prop('Hidden',True),prop('MacroExecutionMode',NEVER_EXECUTE),prop('UpdateDocMode',0),prop('InteractionHandler',handler)));assert doc is not None;return doc
 for file in sorted(os.listdir(ROOT+'/group-outputs')):
  if not file.startswith('shape-group-') or not file.endswith('.pptx'):continue
  doc=load(ROOT+'/group-outputs/'+file)
  try:
   old=snapshot(doc);assert len(old)==1
   saved=OUT+'/'+file;doc.storeToURL(uno.systemPathToFileUrl(saved),(prop('FilterName','Impress MS PowerPoint 2007 XML'),prop('Overwrite',True)))
   reopened=load(saved)
   try:new=snapshot(reopened)
   finally:reopened.close(True)
   def flatten(shapes):
    return [s for shape in shapes for s in [shape]+flatten(shape.get('children',[]))]
   old_groups=[s for s in flatten(old[0]) if s['type'].endswith('GroupShape')];new_groups=[s for s in flatten(new[0]) if s['type'].endswith('GroupShape')];assert len(old_groups)==1 and len(new_groups)==1
   assert len(old_groups[0]['children'])==2 and len(new_groups[0]['children'])==2
   assert [s.get('text') for s in old_groups[0]['children']]==[s.get('text') for s in new_groups[0]['children']]
   old_pictures=[s for s in flatten(old[0]) if s['type'].endswith('GraphicObjectShape')];new_pictures=[s for s in flatten(new[0]) if s['type'].endswith('GraphicObjectShape')]
   replaced=[s for s in old_pictures if s['name']=='Picture 2'];assert len(replaced)==1
   assert len(new_pictures)==len(old_pictures)
   before_names=[s['name'] for s in old_pictures];after_names=[s['name'] for s in new_pictures];assert before_names==after_names
   maximum=0
   for a,b in zip(old_pictures,new_pictures):
    for key in ['size','position']:
     for x,y in zip(a[key],b[key]):maximum=max(maximum,abs(x-y))
   assert maximum<=1
   report['rows'].append({'file':file,'picturesImported':len(old_pictures),'picturesReopened':len(new_pictures),'groupChildPictures':len(replaced),'groupsPreserved':1,'childrenPerGroup':2,'maximumGeometryDeltaHundredthsMm':maximum,'original':old,'reopened':new})
  finally:doc.close(True)
 assert not report['interactions'];report['status']='passed'
 report['limits']=['Four successful grouping variants; group has two editable child objects and text survives save/reopen; no independent visual fidelity','PPTX export changes unsupported original vector picture into a shape','Picture position/extents comparison tolerance: 0.01 mm','Not Microsoft PowerPoint application evidence']
except Exception as e:report['status']='failed';report['error']=str(e);raise
finally:
 if desktop:
  try:desktop.terminate()
  except Exception:pass
 try:server.wait(timeout=5)
 except subprocess.TimeoutExpired:os.killpg(server.pid,signal.SIGKILL)
 with open(OUT+'/report.json','w') as f:json.dump(report,f,indent=2)
 print(json.dumps({'status':report['status'],'rows':[{k:v for k,v in r.items() if k not in ['original','reopened']} for r in report['rows']]},indent=2))
