//! One-session contained row edits. No other document/frame is accessible here.
use super::*;
use crate::grids::{GridColumnId,GridKind,MAX_GRID_ROWS};
use crate::paint::{ColorSpec,RgbColor};

#[derive(Clone,Copy)]
struct CellOwner { grid:NodeId,row:NodeId,index:usize }
fn cell_owner(index:&TreeIndex<'_>,cell:NodeId)->Option<CellOwner>{
    let row=index.nodes.get(&cell)?.parent?;
    let NodeKind::GridRow{cells,..}=&index.node(row).ok()?.kind else{return None};
    if cells.len()!=1 || cells[0].id!=cell{return None;}
    let grid=index.nodes.get(&row)?.parent?;
    let NodeKind::Grid{environment:GridKind::Gathered,rows,..}=&index.node(grid).ok()?.kind else{return None};
    Some(CellOwner{grid,row,index:rows.iter().position(|r|r.id==row)?})
}
pub(super) fn inherited_color(node:&Node,target:NodeId,color:RgbColor)->Option<RgbColor>{
    if node.id==target{return Some(color);}
    match &node.kind{
        NodeKind::Row{entries,..}=>{
            let mut current=color;
            for entry in entries{
                if let Some(found)=inherited_color(&entry.node,target,current){return Some(found);}
                if let NodeKind::ColorChange{spec}=&entry.node.kind{current=spec.color;}
            }None
        }
        NodeKind::Color{spec,body,..}=>inherited_color(body,target,spec.color),
        _=>node.children().into_iter().find_map(|child|inherited_color(child,target,color)),
    }
}
pub(super) fn end_color(entries:&[RowEntry],mut color:RgbColor)->RgbColor{
    for entry in entries{if let NodeKind::ColorChange{spec}=&entry.node.kind{color=spec.color;}}
    color
}
impl Builder<'_>{
    pub(super) fn split_line(&mut self)->EngineResult<()> {
        if self.aligned_split_line()? {return Ok(());}
        if self.matrix_split_line()? {return Ok(());}
        let Selection::Range{anchor,focus}=&self.selection else{return Err(EngineError::new(ErrorCode::InvalidPosition,"Place a caret at a formula-row boundary before splitting"));};
        if !anchor.same_location(*focus){return Err(EngineError::new(ErrorCode::InvalidPosition,"Line break requires a collapsed mathematical caret"));}
        let Position::Row{mut node_id,offset,..}=*focus else {
            return Err(EngineError::new(ErrorCode::InvalidPosition,"Leave the text run before splitting a formula line"));
        };
        let mut at=offset as usize;
        let index=TreeIndex::new(&self.tree)?;
        // An edge outside a sole existing gathered group continues that group,
        // instead of wrapping it in a second multiline environment.
        if node_id==self.tree.id {
            if let NodeKind::Row{entries,..}=&self.tree.kind {
                if entries.len()==1 {
                    if let NodeKind::Grid{environment:GridKind::Gathered,rows,..}=&entries[0].node.kind{
                        let row=if at==0 {&rows[0]}else{rows.last().ok_or_else(invalid_position)?};
                        let NodeKind::GridRow{cells,..}=&row.kind else{return Err(invalid_position());};
                        let NodeKind::Row{entries,..}=&cells[0].kind else{return Err(invalid_position());};
                        node_id=cells[0].id;at=if at==0{0}else{entries.len()};
                    }
                }
            }
        }
        let owner=cell_owner(&index,node_id);
        if owner.is_none()&&node_id!=self.tree.id {
            return Err(EngineError::new(ErrorCode::InvalidPosition,"Move outside the fraction, script or other inner structure before breaking the formula line"));
        }
        let NodeKind::Row{entries,trailing,..}=&index.node(node_id)?.kind else{return Err(invalid_position());};
        if at>entries.len(){return Err(invalid_position());}
        let entries=entries.clone();let trailing=trailing.clone();
        if let Some(owner)=owner {
            let NodeKind::Grid{rows,..}=&index.node(owner.grid)?.kind else{return Err(invalid_position());};
            if rows.len()>=MAX_GRID_ROWS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Contained row budget exhausted"));}
        }
        let inherited=inherited_color(&self.tree,node_id,RgbColor{r:0,g:0,b:0}).ok_or_else(invalid_position)?;
        let active=end_color(&entries[..at],inherited);
        let mut right_entries=entries[at..].to_vec();let prefix=usize::from(active!=inherited);
        if prefix!=0 {
            let node=self.node(NodeKind::ColorChange{spec:ColorSpec::html(active)})?;
            right_entries.insert(0,RowEntry{leading:String::new(),node});
        }
        let right=self.node(NodeKind::Row{grouping:Grouping::Bare,entries:right_entries,trailing})?;
        let right_id=right.id;
        let left_id;
        if let Some(owner)=owner {
            left_id=node_id;
            let mut left=self.tree.find(node_id).ok_or_else(invalid_position)?.clone();left.origin=None;
            left.kind=NodeKind::Row{grouping:Grouping::Bare,entries:entries[..at].to_vec(),trailing:String::new()};
            let NodeKind::GridRow{row_break,..}=&self.tree.find(owner.row).ok_or_else(invalid_position)?.kind else{return Err(invalid_position());};
            let right_row=self.node(NodeKind::GridRow{cells:vec![right],separators:Vec::new(),row_break:row_break.clone()})?;
            self.update(owner.grid,move|grid|{
                let NodeKind::Grid{rows,..}=&mut grid.kind else{return Err(invalid_position());};
                let row=Arc::make_mut(&mut rows[owner.index]);row.origin=None;
                row.kind=NodeKind::GridRow{cells:vec![Arc::new(left)],separators:Vec::new(),row_break:"\\\\".into()};
                rows.insert(owner.index+1,right_row);Ok(())
            })?;
        }else{
            let left=self.row(Grouping::Bare,entries[..at].to_vec())?;left_id=left.id;
            let left_row=self.node(NodeKind::GridRow{cells:vec![left],separators:Vec::new(),row_break:"\\\\".into()})?;
            let right_row=self.node(NodeKind::GridRow{cells:vec![right],separators:Vec::new(),row_break:String::new()})?;
            let column=GridColumnId::allocate(self.ids)?;
            let grid=self.node(NodeKind::Grid{environment:GridKind::Gathered,columns:vec![column],array_spec:None,
                before_begin_name:String::new(),before_end_name:String::new(),rows:vec![left_row,right_row]})?;
            self.update(node_id,move|root|{
                root.kind=NodeKind::Row{grouping:Grouping::Bare,entries:vec![RowEntry{leading:String::new(),node:grid}],trailing:String::new()};Ok(())
            })?;
        }
        self.overrides.push(PositionStep::RowSplit{node_id,left_id,right_id,at:at as u32,right_offset:prefix as u32});
        self.selection=Selection::caret(Position::row(right_id,prefix));Ok(())
    }

    /// true means the current grid owns the boundary, including first/last-row
    /// no-ops. The caller must not fall through to an outer document deletion.
    pub(super) fn merge_line(&mut self,direction:Horizontal)->EngineResult<bool>{
        if self.aligned_merge_line(direction)? {return Ok(true);}
        if self.matrix_delete_boundary(direction)? {return Ok(true);}
        let Selection::Range{anchor,focus}=&self.selection else{return Ok(false)};
        if !anchor.same_location(*focus){return Ok(false)}
        let Position::Row{node_id,offset,..}=*focus else{return Ok(false)};
        let index=TreeIndex::new(&self.tree)?;let Some(owner)=cell_owner(&index,node_id)else{return Ok(false)};
        let NodeKind::Row{entries,..}=&index.node(node_id)?.kind else{return Err(invalid_position());};
        if (direction==Horizontal::Left&&offset!=0)||(direction==Horizontal::Right&&offset as usize!=entries.len()){return Ok(false)}
        let NodeKind::Grid{rows,..}=&index.node(owner.grid)?.kind else{return Err(invalid_position());};
        if (direction==Horizontal::Left&&owner.index==0)||(direction==Horizontal::Right&&owner.index+1==rows.len()){return Ok(true)}
        let left_index=if direction==Horizontal::Left{owner.index-1}else{owner.index};
        let left_row=rows[left_index].as_ref().clone();let right_row=rows[left_index+1].as_ref().clone();let row_count=rows.len();
        let NodeKind::GridRow{cells:left_cells,..}=&left_row.kind else{return Err(invalid_position());};
        let NodeKind::GridRow{cells:right_cells,row_break,..}=&right_row.kind else{return Err(invalid_position());};
        let left_cell=left_cells[0].as_ref().clone();let right_cell=right_cells[0].as_ref().clone();
        let (NodeKind::Row{entries:a,trailing:ta,..},NodeKind::Row{entries:b,trailing:tb,..})=(&left_cell.kind,&right_cell.kind)else{return Err(invalid_position());};
        let join=a.len();let mut joined=a.clone();let mut following=b.clone();let mut trivia=ta.clone();
        if crate::model::ends_in_line_comment(&trivia){trivia.push('\n');}
        let inherited=inherited_color(&self.tree,left_cell.id,RgbColor{r:0,g:0,b:0}).ok_or_else(invalid_position)?;
        let reset=end_color(a,inherited)!=inherited;
        if reset {
            following.insert(0,RowEntry{leading:String::new(),node:self.node(NodeKind::ColorChange{spec:ColorSpec::html(inherited)})?});
        }
        let trailing=if let Some(first)=following.first_mut(){first.leading.insert_str(0,&trivia);tb.clone()}else{format!("{trivia}{tb}")};
        joined.extend(following);
        let left_id=left_cell.id;let right_id=right_cell.id;let removed_row=right_row.id;
        // Merging the last two lines of a gathered group that is the whole
        // formula dissolves the group again instead of leaving a one-line
        // \begin{gathered}..\end{gathered} behind.
        let root_id=self.tree.id;
        let dissolve=row_count==2&&matches!(&self.tree.kind,NodeKind::Row{entries,..} if entries.len()==1&&entries[0].node.id==owner.grid);
        if dissolve {
            let count=joined.len()as u32;
            self.update(root_id,move|root|{
                let NodeKind::Row{entries,trailing:root_trailing,..}=&mut root.kind else{return Err(invalid_position());};
                let leading=std::mem::take(&mut entries[0].leading);
                let mut joined=joined;
                let trailing=match joined.first_mut(){Some(first)=>{first.leading.insert_str(0,&leading);format!("{trailing}{root_trailing}")}None=>format!("{leading}{trailing}{root_trailing}")};
                *entries=joined;*root_trailing=trailing;Ok(())
            })?;
            self.overrides.push(PositionStep::RowSplice{node_id:root_id,start:0,removed:1,inserted:count});
            self.overrides.push(PositionStep::RowOffset{node_id:left_id,target_id:root_id,offset:0});
            self.overrides.push(PositionStep::RowOffset{node_id:right_id,target_id:root_id,offset:(join+usize::from(reset))as u32});
            self.selection=Selection::caret(Position::row(root_id,join));return Ok(true);
        }
        let mut left=left_cell;left.origin=None;left.kind=NodeKind::Row{grouping:Grouping::Bare,entries:joined,trailing};
        let row_break=row_break.clone();
        self.update(owner.grid,move|grid|{
            let NodeKind::Grid{rows,..}=&mut grid.kind else{return Err(invalid_position());};
            let row=Arc::make_mut(&mut rows[left_index]);row.origin=None;
            row.kind=NodeKind::GridRow{cells:vec![Arc::new(left)],separators:Vec::new(),row_break};
            rows.remove(left_index+1);Ok(())
        })?;
        self.overrides.push(PositionStep::RowOffset{node_id:left_id,target_id:left_id,offset:0});
        self.overrides.push(PositionStep::RowOffset{node_id:right_id,target_id:left_id,offset:(join+usize::from(reset))as u32});
        self.overrides.push(PositionStep::Removed{node_id:removed_row,fallback:Position::row(left_id,join)});
        self.selection=Selection::caret(Position::row(left_id,join));Ok(true)
    }
}
