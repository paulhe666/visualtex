//! Aligned row splits/merges preserve column ownership and actual math nodes.
//! No source regex, whole-document normalization or neighboring host access.
use super::*;
use super::matrix::nearest_matrix;
use super::lines::{inherited_color,end_color};
use crate::paint::{ColorSpec,RgbColor};

impl Builder<'_> {
    pub(super) fn aligned_split_line(&mut self)->EngineResult<bool>{
        let Some(location)=nearest_matrix(&self.tree,&self.selection)? else{return Ok(false)};
        let NodeKind::Grid{environment,..}=&self.tree.find(location.grid).ok_or_else(invalid_position)?.kind else{return Err(invalid_position())};
        if !environment.is_aligned(){return Ok(false);}
        let (columns,mut rows)=self.matrix_parts(location.grid)?;
        if rows.len()>=crate::grids::MAX_GRID_ROWS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Aligned row budget exhausted"));}
        let focus=self.selection.focus().ok_or_else(invalid_position)?;
        // Nested structures are kept whole. An explicit row insertion is safer
        // than slicing a fraction, script or text run at a source offset.
        if focus.node_id()!=location.cell {
            self.insert_matrix_row_at(location,true)?;
            let index=TreeIndex::new(&self.tree)?;
            self.selection=Selection::caret(Position::row(index.cell_at(location.grid,location.row+1,0)?,0));
            return Ok(true);
        }
        let NodeKind::GridRow{cells,separators,row_break}=&rows[location.row].kind else{return Err(invalid_position())};
        let mut left_cells=cells.clone();let separators=separators.clone();let old_break=row_break.clone();
        let mut right_cells=Vec::with_capacity(columns.len());
        let at=focus.offset();let active_len=self.row_entries(location.cell)?.len();
        if at>active_len{return Err(invalid_position());}
        let end_of_row=location.column+1==columns.len()&&at==active_len;
        let mut caret_prefix=0;
        for c in 0..columns.len(){
            let original=left_cells[c].clone();
            if c<location.column {right_cells.push(self.row(Grouping::Bare,Vec::new())?);continue;}
            let NodeKind::Row{entries,trailing,..}=&original.kind else{return Err(invalid_position())};
            let split=if c==location.column{at}else{0};
            let inherited=inherited_color(&self.tree,original.id,RgbColor{r:0,g:0,b:0}).ok_or_else(invalid_position)?;
            let active=end_color(&entries[..split],inherited);let prefix=usize::from(active!=inherited);
            let mut right_entries=entries[split..].to_vec();
            if prefix!=0 {right_entries.insert(0,RowEntry{leading:String::new(),node:self.node(NodeKind::ColorChange{spec:ColorSpec::html(active)})?});}
            let right=self.node(NodeKind::Row{grouping:Grouping::Bare,entries:right_entries,trailing:trailing.clone()})?;
            self.overrides.push(PositionStep::RowSplit{node_id:original.id,left_id:original.id,right_id:right.id,at:split as u32,right_offset:prefix as u32});
            left_cells[c]=Arc::new(Node{id:original.id,origin:None,kind:NodeKind::Row{
                grouping:Grouping::Bare,entries:entries[..split].to_vec(),trailing:String::new()}});
            if c==location.column{caret_prefix=prefix;}
            right_cells.push(right);
        }
        let target_column=if end_of_row{0}else{location.column};let target=right_cells[target_column].id;
        let right_row=self.node(NodeKind::GridRow{cells:right_cells,separators:separators.clone(),row_break:old_break})?;
        let left=Arc::make_mut(&mut rows[location.row]);left.origin=None;
        left.kind=NodeKind::GridRow{cells:left_cells,separators,row_break:"\\\\".into()};
        rows.insert(location.row+1,right_row);
        self.update(location.grid,move|grid|{let NodeKind::Grid{rows:stored,..}=&mut grid.kind else{return Err(invalid_position())};*stored=rows;Ok(())})?;
        self.selection=Selection::caret(Position::row(target,if end_of_row{0}else{caret_prefix}));Ok(true)
    }

    pub(super) fn aligned_merge_line(&mut self,direction:Horizontal)->EngineResult<bool>{
        let Some(location)=nearest_matrix(&self.tree,&self.selection)? else{return Ok(false)};
        let NodeKind::Grid{environment,..}=&self.tree.find(location.grid).ok_or_else(invalid_position)?.kind else{return Err(invalid_position())};
        if !environment.is_aligned(){return Ok(false);}
        let focus=self.selection.focus().ok_or_else(invalid_position)?;
        if focus.node_id()!=location.cell{return Ok(false);}
        let (columns,mut rows)=self.matrix_parts(location.grid)?;
        let at_boundary=match direction {
            Horizontal::Left=>location.column==0&&focus.offset()==0,
            Horizontal::Right=>location.column+1==columns.len()&&focus.offset()==self.row_entries(location.cell)?.len(),
        };
        if !at_boundary{return Ok(false);}
        if (direction==Horizontal::Left&&location.row==0)||(direction==Horizontal::Right&&location.row+1==rows.len()){return Ok(true);}
        let at=if direction==Horizontal::Left{location.row-1}else{location.row};
        let NodeKind::GridRow{cells:left_cells,separators,..}=&rows[at].kind else{return Err(invalid_position())};
        let NodeKind::GridRow{cells:right_cells,row_break,..}=&rows[at+1].kind else{return Err(invalid_position())};
        let left_cells=left_cells.clone();let right_cells=right_cells.clone();let separators=separators.clone();let row_break=row_break.clone();
        let removed_row=rows[at+1].id;let mut merged=Vec::with_capacity(columns.len());let mut caret=None;
        for c in 0..columns.len(){
            let left=&left_cells[c];let right=&right_cells[c];
            let (NodeKind::Row{entries:a,trailing:ta,..},NodeKind::Row{entries:b,trailing:tb,..})=(&left.kind,&right.kind)else{return Err(invalid_position())};
            let join=a.len();let mut entries=a.clone();let mut following=b.clone();let mut trivia=ta.clone();
            if crate::model::ends_in_line_comment(&trivia){trivia.push('\n');}
            let inherited=inherited_color(&self.tree,left.id,RgbColor{r:0,g:0,b:0}).ok_or_else(invalid_position)?;
            let reset=usize::from(end_color(a,inherited)!=inherited);
            if reset!=0 {following.insert(0,RowEntry{leading:String::new(),node:self.node(NodeKind::ColorChange{spec:ColorSpec::html(inherited)})?});}
            let trailing=if let Some(first)=following.first_mut(){first.leading.insert_str(0,&trivia);tb.clone()}else{format!("{trivia}{tb}")};
            entries.extend(following);
            self.overrides.push(PositionStep::RowOffset{node_id:left.id,target_id:left.id,offset:0});
            self.overrides.push(PositionStep::RowOffset{node_id:right.id,target_id:left.id,offset:(join+reset)as u32});
            if c==location.column{caret=Some(Position::row(left.id,join));}
            merged.push(Arc::new(Node{id:left.id,origin:None,kind:NodeKind::Row{grouping:Grouping::Bare,entries,trailing}}));
        }
        let caret=caret.ok_or_else(invalid_position)?;
        let left=Arc::make_mut(&mut rows[at]);left.origin=None;left.kind=NodeKind::GridRow{cells:merged,separators,row_break};
        rows.remove(at+1);
        self.update(location.grid,move|grid|{let NodeKind::Grid{rows:stored,..}=&mut grid.kind else{return Err(invalid_position())};*stored=rows;Ok(())})?;
        self.overrides.push(PositionStep::Removed{node_id:removed_row,fallback:caret});
        self.selection=Selection::caret(caret);Ok(true)
    }
}
