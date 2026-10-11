//! Structure-aware vertical landing. Region preorder is emitted by flatten;
//! source lengths, glyph order and DOM rectangles are never used as anchors.
use std::collections::BTreeSet;
use crate::positions::Position;
use crate::scene::{CaretStop, LayoutSnapshot};
use crate::session::NodeId;

fn node_id(position: Position) -> NodeId {
    match position { Position::Row {node_id,..}|Position::Text {node_id,..}=>node_id }
}
fn subtree_end(scene: &LayoutSnapshot, start: usize) -> usize {
    let depth=scene.regions[start].depth;
    (start+1..scene.regions.len()).find(|i|scene.regions[*i].depth<=depth)
        .unwrap_or(scene.regions.len())
}

fn grid_target(scene:&LayoutSnapshot,grid:usize,current:usize,down:bool,x:f64)->Option<Position>{
    let end=subtree_end(scene,grid);
    let rows:Vec<_>=(grid+1..end).filter(|i|scene.regions[*i].depth==scene.regions[grid].depth+1&&scene.regions[*i].role=="gridRow").collect();
    let row=rows.iter().position(|&r|current>=r&&current<subtree_end(scene,r))?;
    let target=if down {rows.get(row+1).copied()} else {row.checked_sub(1).and_then(|i|rows.get(i).copied())}?;
    let direct_cells=|owner:usize| (owner+1..subtree_end(scene,owner))
        .filter(|i|scene.regions[*i].depth==scene.regions[owner].depth+1&&scene.regions[*i].role=="row").collect::<Vec<_>>();
    let source_cells=direct_cells(rows[row]);
    let column=source_cells.iter().position(|&cell|current>=cell&&current<subtree_end(scene,cell))?;
    let cell=*direct_cells(target).get(column)?;
    let ids:BTreeSet<_>=scene.regions[cell..subtree_end(scene,cell)].iter().map(|r|r.node_id).collect();
    let baseline=scene.carets.iter().find(|c|node_id(c.position)==scene.regions[cell].node_id)?.baseline;
    scene.carets.iter().filter(|c|ids.contains(&node_id(c.position))&&(c.baseline-baseline).abs()<1e-9)
        .min_by(|a,b|(a.x-x).abs().total_cmp(&(b.x-x).abs()).then(b.depth.cmp(&a.depth))).map(|c|c.position)
}

/// Prefer a sibling mathematical argument before unrelated geometric neighbors.
/// The root row is not an intermediate editable line between fraction arguments
/// or the upper/lower limits of an operator.
pub(crate) fn fraction_target(scene:&LayoutSnapshot, from:Position, down:bool, desired_x:f64)->Option<Position> {
    let index=scene.regions.iter().position(|r|r.node_id==node_id(from))?;
    let mut ancestors=Vec::new();
    for i in 0..=index {
        let depth=scene.regions[i].depth;
        while ancestors.last().is_some_and(|j:&usize|scene.regions[*j].depth>=depth) {ancestors.pop();}
        ancestors.push(i);
    }
    for &parent in ancestors.iter().rev() {
        if scene.regions[parent].role=="grid" {
            if let Some(position)=grid_target(scene,parent,index,down,desired_x){return Some(position);}
            continue;
        }
        if !matches!(scene.regions[parent].role.as_str(),"fraction"|"largeOperator"|"generalOperator"|"customSymbol"|"scripts"|"annotation") {continue;}
        let end=subtree_end(scene,parent);
        let direct:Vec<_>=(parent+1..end).filter(|i|scene.regions[*i].depth==scene.regions[parent].depth+1).collect();
        // Base/superscript/subscript and label/base move by actual baseline.
        let general=matches!(scene.regions[parent].role.as_str(),"generalOperator"|"scripts"|"annotation");
        if (general && direct.len()<2) || (!general && direct.len()!=2) {continue;}
        // Use the children's actual baselines so authored source order never
        // decides the direction of navigation. This also covers side scripts.
        let baseline=|region:usize|scene.carets.iter().find(|c|node_id(c.position)==scene.regions[region].node_id).map(|c|c.baseline);
        let contains=|region:usize|index>=region&&index<subtree_end(scene,region);
        let target=if general {
            let current=direct.iter().copied().find(|&i|contains(i))?;
            let origin=baseline(current)?;
            direct.iter().copied().filter_map(|i|baseline(i).map(|b|(i,b)))
                .filter(|(i,b)|*i!=current && if down{*b>origin+1e-9}else{*b<origin-1e-9})
                .min_by(|a,b|(a.1-origin).abs().total_cmp(&(b.1-origin).abs())).map(|(i,_)|i)
        } else {
            let (a,b)=(direct[0],direct[1]);let (ab,bb)=(baseline(a)?,baseline(b)?);
            let (upper,lower)=if ab<bb{(a,b)}else{(b,a)};
            if down&&contains(upper){Some(lower)}else if !down&&contains(lower){Some(upper)}else{None}
        };
        let Some(target)=target else {continue;};
        let target_end=subtree_end(scene,target);
        let ids:BTreeSet<_>=scene.regions[target..target_end].iter().map(|r|r.node_id).collect();
        let baseline=scene.carets.iter().find(|c|node_id(c.position)==scene.regions[target].node_id)?.baseline;
        let candidate=scene.carets.iter().filter(|c|ids.contains(&node_id(c.position)) && (c.baseline-baseline).abs()<1e-9)
            .min_by(|a,b| {
                let distance=|c:&CaretStop|(c.x-desired_x).abs();
                distance(a).total_cmp(&distance(b)).then(b.depth.cmp(&a.depth))
            });
        if let Some(candidate)=candidate {return Some(candidate.position);}
    }
    None
}

/// The innermost contained line (grid row) holding a caret, if any.
pub(crate) fn line_of(scene:&LayoutSnapshot, from:Position)->Option<(usize,usize)> {
    let index=scene.regions.iter().position(|r|r.node_id==node_id(from))?;
    let mut best=None;
    for (i,region) in scene.regions.iter().enumerate().take(index+1) {
        if region.role=="gridRow" && index<subtree_end(scene,i) { best=Some((i,subtree_end(scene,i))); }
    }
    best
}
/// Node ids drawn inside a region subtree.
pub(crate) fn subtree_ids(scene:&LayoutSnapshot, start:usize, end:usize)->BTreeSet<NodeId> {
    scene.regions[start..end].iter().map(|r|r.node_id).collect()
}

/// The grid row a point falls into: the deepest grid whose box contains the
/// point, then its row spanning that y (or the vertically nearest row).
pub(crate) fn line_at(scene:&LayoutSnapshot, x:f64, y:f64)->Option<BTreeSet<NodeId>> {
    let grid=(0..scene.regions.len()).filter(|&i|scene.regions[i].role=="grid"&&scene.regions[i].logical.contains(x,y))
        .max_by_key(|&i|scene.regions[i].depth)?;
    let end=subtree_end(scene,grid);let depth=scene.regions[grid].depth;
    let row=(grid+1..end).filter(|&i|scene.regions[i].depth==depth+1&&scene.regions[i].role=="gridRow")
        .min_by(|&a,&b|{
            let distance=|i:usize|{let r=scene.regions[i].logical;if y<r.y_min{r.y_min-y}else if y>r.y_max{y-r.y_max}else{0.0}};
            distance(a).total_cmp(&distance(b))
        })?;
    let ids=subtree_ids(scene,row,subtree_end(scene,row));
    scene.carets.iter().any(|c|ids.contains(&c.position.node_id())).then_some(ids)
}
