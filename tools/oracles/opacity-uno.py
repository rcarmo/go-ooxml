#!/usr/bin/python3
"""Optional test-only LibreOffice shape opacity and picture transparency import/save/reopen oracle."""
import json, os, signal, subprocess, sys, time, uuid
from project_tmp import oracle_profile
import uno, unohelper
from com.sun.star.task import XInteractionHandler
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
ROOT=os.path.abspath(sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(__file__),'../../artifacts/graphics'));OUT=ROOT+'/opacity-uno';os.makedirs(OUT,exist_ok=True)
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
  properties=shape.getPropertySetInfo()
  for key in ['RotateAngle','GraphicCrop','GraphicIsMirrored','Transparency','FillTransparence']:
   if properties.hasPropertyByName(key):
    v=shape.getPropertyValue(key)
    if key=='GraphicCrop':v={k:getattr(v,k) for k in ['Top','Bottom','Left','Right']}
    row[key]=v
  if properties.hasPropertyByName('FillStyle'):row['FillStyle']=str(shape.getPropertyValue('FillStyle'))
  try:row['text']=shape.getString()
  except Exception:pass
  if shape.supportsService('com.sun.star.drawing.GroupShape'):row['children']=[scan(shape.getByIndex(i)) for i in range(shape.getCount())]
  return row
 return [[scan(pages.getByIndex(i).getByIndex(j)) for j in range(pages.getByIndex(i).getCount())] for i in range(pages.getCount())]
pipe='go_pictures_'+uuid.uuid4().hex
profile=oracle_profile()
server=subprocess.Popen(['libreoffice','-env:UserInstallation=file://'+profile,'--headless','--nologo','--nodefault','--nofirststartwizard','--accept=pipe,name='+pipe+';urp;StarOffice.ServiceManager'],stdout=open(OUT+'/server.stdout.log','w'),stderr=open(OUT+'/server.stderr.log','w'),start_new_session=True)
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
 for file in sorted(os.listdir(ROOT+'/opacity-outputs')):
  if not file.startswith('opacity-') or not file.endswith('.pptx') or 'linked picture' in file:continue
  doc=load(ROOT+'/opacity-outputs/'+file)
  try:
   old=snapshot(doc);assert len(old)==1
   saved=OUT+'/'+file;doc.storeToURL(uno.systemPathToFileUrl(saved),(prop('FilterName','Impress MS PowerPoint 2007 XML'),prop('Overwrite',True)))
   reopened=load(saved)
   try:new=snapshot(reopened)
   finally:reopened.close(True)
   def flatten(shapes):
    return [s for shape in shapes for s in [shape]+flatten(shape.get('children',[]))]
   old_pictures=[s for s in flatten(old[0]) if s['type'].endswith('GraphicObjectShape')];new_pictures=[s for s in flatten(new[0]) if s['type'].endswith('GraphicObjectShape')]
   selected=[s for s in flatten(old[0]) if s['name']=='Gradient target'] if file in ['opacity-theme fill.pptx','opacity-transparent fill.pptx','opacity-opaque no-op.pptx','opacity-existing alpha.pptx','opacity-RGB fill.pptx'] else [s for s in old_pictures if s['name']=='Picture 2'];assert len(selected)==1
   counterpart=next(s for s in flatten(new[0]) if s['name']==selected[0]['name'])
   key='FillTransparence' if selected[0]['name']=='Gradient target' else 'Transparency'
   if selected[0].get(key)!=counterpart.get(key):
    assert key=='FillTransparence' and selected[0][key]==100 and 'NONE' in counterpart.get('FillStyle',''),(file,key,selected[0].get(key),counterpart.get(key),counterpart.get('FillStyle'))
    report.setdefault('normalisations',[]).append({'file':file,'change':'100% transparent solid fill exported as noFill; scalar changes from100 to0 but effective fill is absent'})
   expected={'theme fill':50,'transparent fill':100,'opaque no-op':0,'existing alpha':25,'RGB fill':75,'existing picture':25,'picture no-op':65,'opaque picture':0,'transparent picture':100,'absent picture effect':50,'grouped picture':25}
   case=file[len('opacity-'):-len('.pptx')];assert selected[0][key]==expected[case], 'alpha import literal differs'
   assert len(new_pictures)==len(old_pictures)
   before_names=[s['name'] for s in old_pictures];after_names=[s['name'] for s in new_pictures];assert before_names==after_names
   maximum=0
   for a,b in zip(old_pictures,new_pictures):
    for key in ['size','position']:
     for x,y in zip(a[key],b[key]):maximum=max(maximum,abs(x-y))
    for key in ['RotateAngle','GraphicCrop','GraphicIsMirrored']:assert a.get(key)==b.get(key),key+' changed on reopen'
   assert maximum<=1
   report['rows'].append({'file':file,'picturesImported':len(old_pictures),'picturesReopened':len(new_pictures),'selectedPictures':len(selected),'maximumGeometryDeltaHundredthsMm':maximum,'original':old,'reopened':new})
  finally:doc.close(True)
 assert not report['interactions'];report['status']='passed'
 report['limits']=['Eleven shape/picture alpha outputs; linked case excluded to avoid fetching; literal UNO transparency and reopened scalar checked; fully transparent fill may normalise to noFill; no visual fidelity','PPTX export changes unsupported original vector picture into a shape','Picture position/extents comparison tolerance: 0.01 mm','Not Microsoft PowerPoint application evidence']
except Exception as e:report['status']='failed';report['error']=str(e);raise
finally:
 if desktop:
  try:desktop.terminate()
  except Exception:pass
 try:server.wait(timeout=5)
 except subprocess.TimeoutExpired:os.killpg(server.pid,signal.SIGKILL)
 with open(OUT+'/report.json','w') as f:json.dump(report,f,indent=2)
 print(json.dumps({'status':report['status'],'rows':[{k:v for k,v in r.items() if k not in ['original','reopened']} for r in report['rows']]},indent=2))
