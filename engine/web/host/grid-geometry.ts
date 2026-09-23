// Alignment metadata is part of the same immutable scene as glyph geometry.
// Reject malformed anchor data before replacing a previously valid SVG.
import type { LayoutSnapshot } from './layout';
const id=(value:unknown):value is number=>typeof value==='number'&&Number.isInteger(value)&&value>0&&value<=0xffffffff;
export function validateGridGeometry(snapshot:LayoutSnapshot):void {
  const gridRegions=snapshot.regions.filter(region=>region.role==='grid');
  if(snapshot.gridColumns===undefined&&gridRegions.length===0)return;
  if(!Array.isArray(snapshot.gridColumns)||snapshot.gridColumns.length>8192||snapshot.regions.length>8192)throw new Error('Invalid grid column collection');
  const nodes=new Set<number>();
  for(const region of snapshot.regions){
    if(!id(region.nodeId)||nodes.has(region.nodeId))throw new Error('Duplicate or invalid mathematical region identity');
    nodes.add(region.nodeId);
  }
  const owners=new Map(gridRegions.map(region=>[region.nodeId,region]));
  const used=new Set<number>();const expected=new Map<number,number>();const previousRight=new Map<number,number>();
  for(const anchor of snapshot.gridColumns){
    if(!anchor||typeof anchor!=='object'||!id(anchor.gridId)||!id(anchor.columnId)
      ||used.has(anchor.columnId)||!Number.isInteger(anchor.index)||anchor.index<0||anchor.index>=64
      ||anchor.index!==(expected.get(anchor.gridId)??0)
      ||![anchor.x,anchor.width].every(v=>typeof v==='number'&&Number.isFinite(v))
      ||anchor.width<0||Math.abs(anchor.x)>1e6||anchor.width>1e6)throw new Error('Invalid grid column anchor');
    const owner=owners.get(anchor.gridId);
    if(!owner||![owner.logical.xMin,owner.logical.xMax].every(Number.isFinite)
      ||owner.logical.xMax<owner.logical.xMin||nodes.has(anchor.columnId))throw new Error('Invalid grid column owner');
    const right=anchor.x+anchor.width;
    const epsilon=1e-9*Math.max(1,owner.logical.xMax-owner.logical.xMin,Math.abs(anchor.x),Math.abs(right));
    if(anchor.x<owner.logical.xMin-epsilon||right>owner.logical.xMax+epsilon)throw new Error('Grid column is outside its owner');
    const previous=previousRight.get(anchor.gridId);
    if(previous!==undefined&&anchor.x<previous-epsilon)throw new Error('Grid columns overlap or are out of order');
    used.add(anchor.columnId);expected.set(anchor.gridId,anchor.index+1);previousRight.set(anchor.gridId,right);
  }
  if(expected.size!==owners.size)throw new Error('A mathematical grid is missing its column anchors');
}
