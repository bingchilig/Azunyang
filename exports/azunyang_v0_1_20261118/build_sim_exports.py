"""Build and validate simulation artifacts from measured Fusion metadata only."""
from pathlib import Path
import json, math, xml.etree.ElementTree as ET
import numpy as np
import trimesh
import mujoco

ROOT = Path(__file__).resolve().parent
data = json.loads((ROOT / 'fusion_model.json').read_text(encoding='utf-8'))
assert data['stage'] == 'F_export'
assert all(not p['B'] for p in data['poses'])
def mat(values):
    m = np.array(values, float).reshape(4, 4)
    m[:3, 3] *= .01
    return m
def text(a): return ' '.join(f'{float(v):.12g}' for v in np.ravel(a))
def rpy(r):
    p = math.asin(np.clip(-r[2, 0], -1, 1))
    return [math.atan2(r[2, 1], r[2, 2]), p, math.atan2(r[1, 0], r[0, 0])]
def skew(v):
    x,y,z=v
    return np.array([[0,-z,y],[z,0,-x],[-y,x,0]])
def origin(parent, xyz=(0,0,0), rotation=None):
    return ET.SubElement(parent, 'origin', xyz=text(xyz), rpy=text([0,0,0] if rotation is None else rpy(rotation)))
def tensor(p):
    c = np.array(p['com_cm']) * .01
    mass = p['mass_kg']
    raw = np.array(p['inertia_origin_kgcm2'][1:]) * 1e-4
    target = np.sort(np.array(p['principal_kgcm2'][1:]) * 1e-4)
    best = None
    for sign in [1, -1]:
        a,b,d,xy,yz,xz=raw
        i=np.array([[a,sign*xy,sign*xz],[sign*xy,b,sign*yz],[sign*xz,sign*yz,d]])
        for shift in [True, False]:
            q=i-mass*(np.dot(c,c)*np.eye(3)-np.outer(c,c)) if shift else i
            eig=np.linalg.eigvalsh(q)
            err=np.linalg.norm(eig-target)/max(np.linalg.norm(target),1e-15)
            if eig.min()>0 and (best is None or err<best[0]):best=(err,q,sign,shift)
    assert best is not None and best[0]<.02, ('Inertia convention validation',p,best)
    return mass,c,best[1],{'product_sign':best[2],'origin_shift':best[3],'principal_relative_error':best[0]}

neutral=data['neutral']
frames={'base_link':np.zeros(3)}
joint_data={}
for j in data['joints']:
    m0,m1=mat(j['child_zero']),mat(j['child_1deg'])
    rd=m1[:3,:3]@m0[:3,:3].T
    candidates=[]
    for index,g in enumerate(j['origins']):
        o=j['child'] if index==0 else j['parent']
        parent_m=mat(neutral[o]['matrix'])
        if g.get('native_transform'):
            candidates.append((parent_m@mat(g['native_transform']))[:3,3])
        ng=np.array(g['native_geometry_origin_cm'])*.01
        off=np.array(g['offsets_cm'])*.01
        candidates.append((parent_m@np.r_[ng+off,1])[:3])
        candidates.append((parent_m@np.r_[ng,1])[:3])
        if g.get('proxy_transform'):candidates.append(mat(g['proxy_transform'])[:3,3])
    rhs=m1[:3,3]-rd@m0[:3,3]
    valid=[p for p in candidates if np.linalg.norm((np.eye(3)-rd)@p-rhs)<1e-6]
    assert valid, ('No measured joint origin agrees with perturbation',j['name'])
    pivot=valid[0]
    axis=np.array(j['axis_world']);axis/=np.linalg.norm(axis)
    name=j['name'].replace('_joint','_link')
    frames[name]=pivot
    joint_data[name]=(j,axis)

mesh_keys={m['component']:m['key'] for m in data['meshes']}
mesh_checks=[]
for spec in data['meshes']:
    mesh=trimesh.load_mesh(spec['file'], process=True)
    expected_volume=spec['physical']['volume_cm3']*1e-6
    guess=(expected_volume/abs(mesh.volume))**(1/3)
    scale=min([.001,.01,1,.0254],key=lambda s:abs(math.log(s/guess)))
    assert abs((abs(mesh.volume)*scale**3)/expected_volume-1)<.08, ('STL unit check',spec['key'])
    mesh.apply_scale(scale)
    expected=np.array(spec['bounds_cm'])*.01
    candidates=[('native',np.eye(4))]
    candidates.extend((n,np.linalg.inv(mat(v['matrix']))) for n,v in neutral.items() if v['component']==spec['component'])
    tests=[]
    for name,t in candidates:
        trial=mesh.copy();trial.apply_transform(t)
        tests.append((float(abs(trial.bounds-expected).max()),name,trial))
    err,source,mesh=min(tests,key=lambda x:x[0])
    assert err<.0008, ('STL frame/bounds check',spec['key'],err)
    mesh.export(ROOT/'meshes'/f"{spec['key']}.stl")
    mesh_checks.append({'mesh':spec['key'],'source_to_m':scale,'frame_conversion':source,'bounds_error_m':err})

groups={'base_link':['base_link:1','ACT_ref:2','ACT_ref:5','ACT_ref:8','ACT_ref:11']}
for k,leg in enumerate(['FL','FR','RL','RR']):
    n=k+1
    groups[f'{leg}_abad_link']=[f'FL_abad_link:{n}',f'ACT_ref:{3+3*k}',f'ACT_ref:{4+3*k}']
    groups[f'{leg}_hip_link']=[f'FL_hip_link:{n}']
    groups[f'{leg}_knee_link']=[f'FL_knee_link:{n}',f'FL_foot:{n}']
robot=ET.Element('robot',name='Azunyang_v0_1')
extension=ET.SubElement(robot,'mujoco')
ET.SubElement(extension,'compiler',meshdir='meshes',strippath='true',discardvisual='false',fusestatic='false',inertiafromgeom='false')
inertia_checks=[]
link_masses={}
def collision(link,name,shape,xyz,size):
    c=ET.SubElement(link,'collision',name=name);origin(c,xyz)
    g=ET.SubElement(c,'geometry')
    if shape=='box':ET.SubElement(g,'box',size=text(size))
    elif shape=='cylinder':ET.SubElement(g,'cylinder',radius=str(size[0]),length=str(size[1]))
    else:ET.SubElement(g,'sphere',radius=str(size[0]))
for lname,parts in groups.items():
    link=ET.SubElement(robot,'link',name=lname)
    masses=[]
    for oname in parts:
        info=neutral[oname];t=mat(info['matrix']);p=info['native_physical']
        mass,c,i,check=tensor(p)
        center=t[:3,:3]@c+t[:3,3]-frames[lname]
        inertia=t[:3,:3]@i@t[:3,:3].T
        masses.append((mass,center,inertia))
        inertia_checks.append({'occurrence':oname,**check})
        v=ET.SubElement(link,'visual',name=oname.replace(':','_'))
        origin(v,t[:3,3]-frames[lname],t[:3,:3])
        geom=ET.SubElement(v,'geometry')
        ET.SubElement(geom,'mesh',filename=f"meshes/{mesh_keys[info['component']]}.stl")
    if lname=='base_link':
        # Explicit budget allowance, not a measured CAD part or material.
        mass=sum(data['parameters'][n]['value'] for n in ['E_wiring_allowance_trial','E_fastener_allowance_trial'])
        assert 0<mass<2
        dims=np.array([.3,.2,.1])
        i=mass/12*np.diag([dims[1]**2+dims[2]**2,dims[0]**2+dims[2]**2,dims[0]**2+dims[1]**2])
        masses.append((mass,np.zeros(3),i))
    mass=sum(p[0] for p in masses)
    center=sum(m*c for m,c,_ in masses)/mass
    inertia=sum(i+m*(np.dot(c-center,c-center)*np.eye(3)-np.outer(c-center,c-center)) for m,c,i in masses)
    eig=np.linalg.eigvalsh(inertia)
    assert eig.min()>0 and eig[2]<=eig[:2].sum()+1e-9
    inertial=ET.SubElement(link,'inertial');origin(inertial,center)
    ET.SubElement(inertial,'mass',value=f'{mass:.12g}')
    ET.SubElement(inertial,'inertia',**{k:f'{inertia[i,j]:.12g}' for k,i,j in [('ixx',0,0),('iyy',1,1),('izz',2,2),('ixy',0,1),('ixz',0,2),('iyz',1,2)]})
    link_masses[lname]=mass
    if lname=='base_link':collision(link,'body_approx','box',[0,0,0],[.284,.2,.1])
    elif '_abad_' in lname:
        hip=frames[lname.replace('abad','hip')]-frames[lname]
        collision(link,'abad_bridge_approx','box',hip/2,[.047,.03,.04])
    elif '_hip_' in lname:collision(link,'thigh_approx','cylinder',[0,0,-.1],[.014,.17])
    else:
        k=['FL','FR','RL','RR'].index(lname[:2])+1
        foot=neutral[f'FL_foot:{k}'];tf=mat(foot['matrix'])
        fc=(tf@np.r_[np.array(foot['native_physical']['com_cm'])*.01,1])[:3]-frames[lname]
        collision(link,'shank_approx','cylinder',[0,fc[1],-.1],[.014,.17])
        collision(link,'foot_approx','sphere',fc,[.015])
for lname,(j,axis) in joint_data.items():
    leg,role,_=lname.split('_')
    parent='base_link' if role=='abad' else f'{leg}_abad_link' if role=='hip' else f'{leg}_hip_link'
    joint=ET.SubElement(robot,'joint',name=j['name'],type='revolute')
    ET.SubElement(joint,'parent',link=parent);ET.SubElement(joint,'child',link=lname)
    origin(joint,frames[lname]-frames[parent]);ET.SubElement(joint,'axis',xyz=text(axis))
    side=1 if leg in ['FL','RL'] else -1
    bounds=(-20,20) if role=='abad' else (0,75) if role=='hip' else (-150,0)
    if role!='abad' and side<0:bounds=(-bounds[1],-bounds[0])
    ET.SubElement(joint,'limit',lower=str(math.radians(bounds[0])),upper=str(math.radians(bounds[1])),effort=str(8.5 if role!='knee' else 12.1125),velocity=str(30.8 if role!='knee' else 20.5333333333))
    ET.SubElement(joint,'dynamics',damping='0.05',friction='0')
ET.indent(robot)
urdf=ROOT/'Azunyang_v0.1.urdf'
ET.ElementTree(robot).write(urdf,encoding='utf-8',xml_declaration=True)
model=mujoco.MjModel.from_xml_path(str(urdf))
assert model.njnt==12
sim=mujoco.MjData(model)
for leg in ['FL','FR','RL','RR']:
    sign=1 if leg in ['FL','RL'] else -1
    for role,q in [('abad',0),('hip',sign*45.6),('knee',-sign*91.1)]:
        jid=mujoco.mj_name2id(model,mujoco.mjtObj.mjOBJ_JOINT,f'{leg}_{role}_joint')
        sim.qpos[model.jnt_qposadr[jid]]=math.radians(q)
mujoco.mj_forward(model,sim)
fk=[]
for lname,(j,_) in joint_data.items():
    n=j['child'];m0=mat(neutral[n]['matrix']);ms=mat(data['standing'][n]['matrix'])
    expected=(ms@np.linalg.inv(m0)@np.r_[frames[lname],1])[:3]
    er=ms[:3,:3]@m0[:3,:3].T
    bid=mujoco.mj_name2id(model,mujoco.mjtObj.mjOBJ_BODY,lname)
    error=float(np.linalg.norm(sim.xpos[bid]-expected))
    rerror=float(np.linalg.norm(sim.xmat[bid].reshape(3,3)-er))
    assert error<1e-5 and rerror<1e-5,(lname,error,rerror)
    fk.append({'link':lname,'position_error_m':error,'rotation_matrix_error':rerror})
mjcf=ROOT/'Azunyang_v0.1.xml'
mujoco.mj_saveLastXML(str(mjcf),model)
tree=ET.parse(mjcf);root=tree.getroot();world=root.find('worldbody');body=world.find('body')
assert body is not None and body.attrib['name']=='base_link'
body.set('pos','0 0 0.43');body.insert(0,ET.Element('freejoint',name='floating_base'))
ET.SubElement(world,'geom',name='ground',type='plane',size='2 2 .1',rgba='.3 .3 .3 1',friction='1 .005 .0001')
act=ET.SubElement(root,'actuator')
for j in data['joints']:
    effort=12.1125 if '_knee_' in j['name'] else 8.5
    ET.SubElement(act,'motor',name=j['name']+'_motor',joint=j['name'],gear='1',ctrllimited='true',ctrlrange=text([-effort,effort]))
ET.indent(tree);tree.write(mjcf,encoding='utf-8',xml_declaration=True)
floating=mujoco.MjModel.from_xml_path(str(mjcf));state=mujoco.MjData(floating)
assert floating.nv==18 and floating.nu==12
for _ in range(10):mujoco.mj_step(floating,state)
assert np.isfinite(state.qpos).all() and np.isfinite(state.qvel).all()
report={'mujoco_version':mujoco.__version__,'urdf_joints':model.njnt,'links':len(groups),'floating_nv':floating.nv,'actuators':floating.nu,'mass_kg_with_estimated_allowances':sum(link_masses.values()),'link_masses_kg':link_masses,'mesh_checks':mesh_checks,'inertia_checks':inertia_checks,'standing_fk':fk,'finite_smoke_steps':10,'holds':['Equivalent motor/electronics masses and inertia; hardware not weighed','500g wiring/fastener allowance distributed as body box','Primitive collision geometry; not CAD clearance proof','Discrete CAD poses do not certify continuous ROM','No actuator/control/gait validation; Isaac runtime import not executed']}
(ROOT/'simulation_validation.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k not in ['inertia_checks','standing_fk','mesh_checks']},indent=2))
