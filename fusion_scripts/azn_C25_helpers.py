import adsk.core, adsk.fusion, json, math
A=adsk.core.Application.get()
D=adsk.fusion.Design.cast(A.activeProduct)
R=D.rootComponent
U=D.unitsManager
V=adsk.core.ValueInput.createByString
P=adsk.core.Point3D.create
OP=adsk.fusion.FeatureOperations
def ev(expr):
    return U.evaluateExpression(str(expr),'mm')
def point(x,y,z='0 mm'):
    return P(ev(x),ev(y),ev(z))
def param(n,e,unit='mm',comment='[Autonomous C2-C5] provisional detail; fabrication validation HOLD'):
    p=D.userParameters.itemByName(n)
    if p: p.expression=e
    else: p=D.userParameters.add(n,V(e),unit,comment)
    return p
def occ(n):
    found=[o for o in R.occurrences if o.component.name==n]
    if found: return found[0]
    o=R.occurrences.addNewComponent(adsk.core.Matrix3D.create()); o.component.name=n; return o
def material(key):
    lib=A.materialLibraries.itemById('C1EEA57C-3F56-45FC-B8CB-A9EC46A9994C')
    if key=='Al': return lib.materials.itemById('PrismMaterial-160')
    if key=='Steel': return lib.materials.itemById('PrismMaterial-018')
    if key=='Rubber': return lib.materials.itemById('PrismMaterial-059')
    name='AZN_'+key+'_TDS_density'
    m=D.materials.itemByName(name)
    if not m:
        m=D.materials.addByCopy(lib.materials.itemById('PrismMaterial-279' if key=='PETG' else 'PrismMaterial-059'),name)
        for prop in m.materialProperties:
            if 'Density' in prop.id and prop.objectType.endswith('FloatProperty'):
                prop.value=1250.0 if key=='PETG' else 1220.0
            if prop.id.startswith('structural_Young_modulus'):
                prop.value=2780000.0 if key=='PETG' else 9800.0
            if prop.id=='structural_Minimum_tensile_strength':
                prop.value=51000.0 if key=='PETG' else 27300.0
    return m
def plane(c,name,axis,offset):
    existing=next((p for p in c.constructionPlanes if p.name==name),None)
    if existing: return existing
    pi=c.constructionPlanes.createInput()
    base=c.xZConstructionPlane if axis=='Y' else c.xYConstructionPlane
    pi.setByOffset(base,V(offset))
    pl=c.constructionPlanes.add(pi); pl.name=name; pl.isLightBulbOn=False; return pl
def sk(c,name,axis,offset):
    s=next((s for s in c.sketches if s.name==name),None)
    if s: return s
    s=c.sketches.add(plane(c,name+'_plane',axis,offset)); s.name=name
    return s
def anchor(s,p,x,y):
    # Coordinates are signed by seed geometry, driving lengths remain positive.
    for coord,expr,orientation in [(p.geometry.x,x,adsk.fusion.DimensionOrientations.HorizontalDimensionOrientation),(p.geometry.y,y,adsk.fusion.DimensionOrientations.VerticalDimensionOrientation)]:
        if abs(coord)<1e-8:
            if orientation==adsk.fusion.DimensionOrientations.HorizontalDimensionOrientation:
                s.geometricConstraints.addVerticalPoints(s.originPoint,p)
            else: s.geometricConstraints.addHorizontalPoints(s.originPoint,p)
        else:
            dim=s.sketchDimensions.addDistanceDimension(s.originPoint,p,orientation,p.geometry)
            dim.parameter.expression='abs('+expr+')'
def circle(s,x,z,dia):
    p=point(x,z)
    prior=next((q for q in s.sketchCurves.sketchCircles if q.centerSketchPoint.geometry.distanceTo(p)<1e-7),None)
    c=s.sketchCurves.sketchCircles.addByCenterRadius(p,ev(dia)/2)
    txt=point('('+x+')+('+dia+')','('+z+')+('+dia+')')
    s.sketchDimensions.addDiameterDimension(c,txt).parameter.expression=dia
    if prior:
        if c.centerSketchPoint != prior.centerSketchPoint:
            s.geometricConstraints.addConcentric(c,prior)
    elif abs(p.x)+abs(p.y)<1e-8:
        if c.centerSketchPoint != s.originPoint:
            s.geometricConstraints.addCoincident(c.centerSketchPoint,s.originPoint)
    else: anchor(s,c.centerSketchPoint,x,z)
    return c
def rect(s,x,z,w,h):
    lines=s.sketchCurves.sketchLines.addTwoPointRectangle(point(x,z),point('('+x+')+('+w+')','('+z+')+('+h+')'))
    for i in [0,2]: s.geometricConstraints.addHorizontal(lines.item(i))
    for i in [1,3]: s.geometricConstraints.addVertical(lines.item(i))
    anchor(s,lines.item(0).startSketchPoint,x,z)
    for i,expr in [(0,w),(1,h)]:
        ln=lines.item(i)
        s.sketchDimensions.addDistanceDimension(ln.startSketchPoint,ln.endSketchPoint,adsk.fusion.DimensionOrientations.AlignedDimensionOrientation,ln.startSketchPoint.geometry).parameter.expression=expr
    return lines
def extrude(c,s,name,t,key='Al',cut=None):
    profiles=adsk.core.ObjectCollection.create()
    # For nested circles/rectangles choose only outer material regions.
    candidates=[p for p in s.profiles]
    if len(candidates)>1:
        target=max(candidates,key=lambda p:(p.profileLoops.count,p.areaProperties().area))
        profiles.add(target)
    else:
        for p in candidates: profiles.add(p)
    ei=c.features.extrudeFeatures.createInput(profiles,OP.CutFeatureOperation if cut else OP.NewBodyFeatureOperation)
    ei.setOneSideExtent(adsk.fusion.DistanceExtentDefinition.create(V(t)),adsk.fusion.ExtentDirections.NegativeExtentDirection if abs(s.transform.getCell(1,2)+1)<1e-7 else adsk.fusion.ExtentDirections.PositiveExtentDirection)
    if cut: ei.participantBodies=cut
    f=c.features.extrudeFeatures.add(ei); f.name=name
    if not cut:
        for b in f.bodies: b.name=name; b.material=material(key)
    s.isLightBulbOn=False
    return f
def ring(c,name,x,z,y,od,ident,t,key='Al'):
    b=next((b for b in c.bRepBodies if b.name==name),None)
    if b:return b
    s=sk(c,name+'_sk','Y',y); circle(s,x,z,od)
    if ident:circle(s,x,z,ident)
    return extrude(c,s,name,t,key).bodies.item(0)
def box(c,name,x,y,z,w,depth,height,key='Al'):
    b=next((b for b in c.bRepBodies if b.name==name),None)
    if b:return b
    s=sk(c,name+'_sk','Z',z); rect(s,x,y,w,depth)
    return extrude(c,s,name,height,key).bodies.item(0)
def jo(c,name,x,y,z):
    old=c.jointOrigins.itemByName(name)
    if old:return old
    j=c.jointOrigins.createInput(adsk.fusion.JointGeometry.createByPoint(c.originConstructionPoint))
    j.offsetX=V(x);j.offsetY=V(y);j.offsetZ=V(z)
    result=c.jointOrigins.add(j);result.name=name;return result
def joint(name,moving,source,target,targetorigin,revolute=True):
    old=R.joints.itemByName(name)
    if old:return old
    p=source.createForAssemblyContext(moving); q=targetorigin.createForAssemblyContext(target)
    assert p.assemblyContext.name==moving.name and q.assemblyContext.name==target.name
    ji=R.joints.createInput(p,q)
    if revolute:ji.setAsRevoluteJointMotion(adsk.fusion.JointDirections.YAxisJointDirection)
    else:ji.setAsRigidJointMotion()
    ji.isFlipped=False
    j=R.joints.add(ji); j.name=name
    print(json.dumps({'joint':name,'one':j.occurrenceOne.name,'two':j.occurrenceTwo.name}))
    return j
def cutcircles(c,name,y,t,holes,bodies):
    if any(f.name==name for f in c.features.extrudeFeatures):return
    s=sk(c,name+'_sk','Y',y)
    for x,z,dia in holes:circle(s,x,z,dia)
    pc=adsk.core.ObjectCollection.create()
    for p in s.profiles: pc.add(p)
    ei=c.features.extrudeFeatures.createInput(pc,OP.CutFeatureOperation)
    ei.setOneSideExtent(adsk.fusion.DistanceExtentDefinition.create(V(t)),adsk.fusion.ExtentDirections.PositiveExtentDirection)
    ei.participantBodies=bodies
    f=c.features.extrudeFeatures.add(ei);f.name=name;s.isLightBulbOn=False
    return f

