import type {MathNode} from './syntax';

type Grid = Extract<MathNode,{kind:'grid'}>;
function children(node:MathNode):MathNode[]{
  switch(node.kind){
    case 'row': return node.entries.map(entry=>entry.node);
    case 'grid': return node.rows;
    case 'gridRow': return node.cells;
    case 'fraction': return [node.numerator,node.denominator];
    case 'radical': return node.degree?[node.degree,node.radicand]:[node.radicand];
    case 'annotation': return [node.annotation,node.base];
    case 'scripts': return [node.base,...node.slots.map(slot=>slot.argument)];
    case 'largeOperator': case 'customSymbol': return node.parts.flatMap(part=>part.kind==='script'?[part.slot.argument]:[]);
    case 'generalOperator': return [node.body,...node.parts.flatMap(part=>part.kind==='script'?[part.slot.argument]:[])];
    case 'macroCall': return node.arguments;
    case 'fence':case 'accent':case 'style':case 'phantom':case 'boxed':case 'color': return [node.body];
    default: return [];
  }
}
/** Read-only shortcut routing. The core independently validates the operation;
 * the host never edits this projection or borrows another host's document. */
export function nearestGrid(root:MathNode,nodeId:number):Grid|null{
  function visit(node:MathNode,owner:Grid|null):Grid|null|undefined{
    const grid=node.kind==='grid'?node:owner;
    if(node.id===nodeId)return grid;
    for(const child of children(node)){const found=visit(child,grid);if(found!==undefined)return found;}
    return undefined;
  }
  return visit(root,null)??null;
}
