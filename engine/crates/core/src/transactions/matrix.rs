//! Matrix/cases edits operate on one nearest grid and never access another session.
//! Shift+Enter inserts an empty row; it never cuts an inner fraction/text cell.
use super::*;
use crate::grids::{GridColumnId, GridKind, MAX_GRID_COLUMNS, MAX_GRID_ROWS};

#[derive(Clone, Copy)]
pub(super) struct MatrixCell { pub(super) grid:NodeId, pub(super) row:usize, pub(super) column:usize, pub(super) cell:NodeId }

pub(super) fn nearest_matrix(tree: &Node, selection: &Selection) -> EngineResult<Option<MatrixCell>> {
    let Selection::Range { anchor, focus } = selection else { return Ok(None) };
    if !anchor.same_location(*focus) { return Ok(None); }
    let index = TreeIndex::new(tree)?;
    let mut current = focus.node_id();
    // A caret at the outer edge of a sole matrix still belongs to that matrix
    // for its row command. Never search an adjacent formula or another frame.
    if current == tree.id {
        if let NodeKind::Row { entries, .. } = &tree.kind {
            if entries.len() == 1 {
                if let NodeKind::Grid { environment, rows, .. } = &entries[0].node.kind {
                    if *environment==GridKind::Gathered{return Ok(None);}
                    let row = if focus.offset() == 0 { 0 } else { rows.len()-1 };
                    let NodeKind::GridRow { cells, .. } = &rows[row].kind else { return Err(invalid_position()) };
                    let column = if focus.offset() == 0 { 0 } else { cells.len()-1 };
                    return Ok(Some(MatrixCell { grid: entries[0].node.id, row, column, cell: cells[column].id }));
                }
            }
        }
    }
    while let Some(parent) = index.nodes.get(&current).and_then(|e|e.parent) {
        if let NodeKind::GridRow { cells, .. } = &index.node(parent)?.kind {
            let grid = index.nodes.get(&parent).and_then(|e|e.parent).ok_or_else(invalid_position)?;
            let NodeKind::Grid { environment, rows, .. } = &index.node(grid)?.kind else { return Err(invalid_position()) };
            // Stop at the nearest environment, including a nested gathered.
            if *environment == GridKind::Gathered { return Ok(None); }
            let row = rows.iter().position(|r|r.id==parent).ok_or_else(invalid_position)?;
            let column = cells.iter().position(|c|c.id==current).ok_or_else(invalid_position)?;
            return Ok(Some(MatrixCell { grid, row, column, cell: current }));
        }
        current = parent;
    }
    Ok(None)
}

fn row_cells(row: &Node) -> EngineResult<&Vec<Arc<Node>>> {
    let NodeKind::GridRow { cells, .. } = &row.kind else { return Err(invalid_position()) };
    Ok(cells)
}
fn repair_breaks(rows: &mut [Arc<Node>]) -> EngineResult<()> {
    let count=rows.len();
    for (i,row) in rows.iter_mut().enumerate() {
        let NodeKind::GridRow { row_break, .. } = &row.kind else { return Err(invalid_position()) };
        let expected=if i+1<count { "\\\\" } else { "" };
        if row_break != expected {
            let row=Arc::make_mut(row);row.origin=None;
            let NodeKind::GridRow { row_break, .. } = &mut row.kind else { unreachable!() };
            *row_break=expected.into();
        }
    }
    Ok(())
}
impl Builder<'_> {
    fn matrix_location(&self) -> EngineResult<MatrixCell> {
        nearest_matrix(&self.tree,&self.selection)?.ok_or_else(||EngineError::new(
            ErrorCode::InvalidPosition,"Matrix command requires a collapsed caret in the nearest matrix"))
    }
    pub(super) fn matrix_parts(&self,id:NodeId)->EngineResult<(Vec<GridColumnId>,Vec<Arc<Node>>)> {
        let NodeKind::Grid { environment, columns, rows, .. } = &self.tree.find(id).ok_or_else(invalid_position)?.kind
            else { return Err(invalid_position()) };
        if *environment==GridKind::Gathered{return Err(invalid_position());}
        Ok((columns.clone(),rows.clone()))
    }
    fn removed_matrix_node(&mut self,node:&Node,fallback:Position) {
        node.visit(&mut |node|self.overrides.push(PositionStep::Removed{node_id:node.id,fallback}));
    }
    pub(super) fn insert_matrix(&mut self,rows:u32,columns:u32)->EngineResult<()> {
        self.insert_cell_environment(GridKind::Matrix,rows,columns)
    }
    pub(super) fn insert_cases(&mut self,rows:u32)->EngineResult<()> {
        self.insert_cell_environment(GridKind::Cases,rows,2)
    }
    pub(super) fn insert_aligned(&mut self,rows:u32,pairs:u32)->EngineResult<()> {
        let columns=pairs.checked_mul(2).ok_or_else(||EngineError::new(ErrorCode::ResourceLimit,"Aligned pair count overflow"))?;
        self.insert_cell_environment(GridKind::Aligned,rows,columns)
    }
    fn insert_cell_environment(&mut self,environment:GridKind,rows:u32,columns:u32)->EngineResult<()> {
        let (rows,columns)=(rows as usize,columns as usize);
        if rows==0 || rows>MAX_GRID_ROWS || columns==0 || columns>MAX_GRID_COLUMNS
            || rows*columns+rows+columns+1>crate::model::MAX_NODES {
            return Err(EngineError::new(ErrorCode::ResourceLimit,"Matrix dimensions exceed the structural budget"));
        }
        let (owner,start,end)=self.ensure_row_range()?;
        let selected=self.row_entries(owner)?[start..end].to_vec();
        let column_ids=(0..columns).map(|_|GridColumnId::allocate(self.ids)).collect::<EngineResult<Vec<_>>>()?;
        let mut grid_rows=Vec::with_capacity(rows);let mut first=None;
        for r in 0..rows {
            let mut cells=Vec::with_capacity(columns);
            for c in 0..columns {
                let cell=self.row(Grouping::Bare,if r==0&&c==0 {selected.clone()}else{Vec::new()})?;
                if first.is_none(){first=Some(cell.id);}
                cells.push(cell);
            }
            grid_rows.push(self.node(NodeKind::GridRow {cells,separators:vec!["&".into();columns-1],
                row_break:if r+1<rows {"\\\\".into()}else{String::new()}})?);
        }
        let grid=self.node(NodeKind::Grid {environment,columns:column_ids,array_spec:None,
            before_begin_name:String::new(),before_end_name:String::new(),rows:grid_rows})?;
        self.splice_row(owner,start,end,vec![RowEntry{leading:String::new(),node:grid}],true)?;
        self.selection=Selection::caret(Position::row(first.ok_or_else(invalid_position)?,selected.len()));
        Ok(())
    }
    pub(super) fn insert_grid_row(&mut self,after:bool)->EngineResult<()> {
        let location=self.matrix_location()?;self.insert_matrix_row_at(location,after)
    }
    pub(super) fn insert_matrix_row_at(&mut self,location:MatrixCell,after:bool)->EngineResult<()> {
        let (columns,mut rows)=self.matrix_parts(location.grid)?;
        if rows.len()>=MAX_GRID_ROWS {return Err(EngineError::new(ErrorCode::ResourceLimit,"Matrix row budget exhausted"));}
        let cells=(0..columns.len()).map(|_|self.row(Grouping::Bare,Vec::new())).collect::<EngineResult<Vec<_>>>()?;
        let target=cells[location.column].id;
        let row=self.node(NodeKind::GridRow {cells,separators:vec!["&".into();columns.len()-1],row_break:String::new()})?;
        rows.insert(location.row+usize::from(after),row);repair_breaks(&mut rows)?;
        self.update(location.grid,move|grid|{
            let NodeKind::Grid{rows:stored,..}=&mut grid.kind else{return Err(invalid_position())};
            *stored=rows;Ok(())
        })?;
        self.selection=Selection::caret(Position::row(target,0));Ok(())
    }
    pub(super) fn delete_grid_row(&mut self)->EngineResult<()> {
        let location=self.matrix_location()?;let (_,mut rows)=self.matrix_parts(location.grid)?;
        if rows.len()==1 {return Err(EngineError::new(ErrorCode::InvalidInput,"A matrix must retain at least one row"));}
        let removed=rows.remove(location.row);let target_row=location.row.min(rows.len()-1);
        let target=row_cells(&rows[target_row])?[location.column].id;let fallback=Position::row(target,0);
        self.removed_matrix_node(&removed,fallback);repair_breaks(&mut rows)?;
        self.update(location.grid,move|grid|{
            let NodeKind::Grid{rows:stored,..}=&mut grid.kind else{return Err(invalid_position())};
            *stored=rows;Ok(())
        })?;
        self.selection=Selection::caret(fallback);Ok(())
    }
    fn require_resizable_columns(&self,grid:NodeId)->EngineResult<()> {
        if !self.tree.find(grid).and_then(|n|match &n.kind{NodeKind::Grid{environment,..}=>Some(environment.is_matrix()),_=>None}).unwrap_or(false) {
            return Err(EngineError::new(ErrorCode::InvalidInput,"Only matrix-family grids accept individual column resizing; Cases/Aligned keep their column schema"));
        }Ok(())
    }
    pub(super) fn insert_grid_column(&mut self,after:bool)->EngineResult<()> {
        let location=self.matrix_location()?;self.require_resizable_columns(location.grid)?;
        let (mut columns,mut rows)=self.matrix_parts(location.grid)?;
        if columns.len()>=MAX_GRID_COLUMNS {return Err(EngineError::new(ErrorCode::ResourceLimit,"Matrix column budget exhausted"));}
        let at=location.column+usize::from(after);columns.insert(at,GridColumnId::allocate(self.ids)?);
        let mut target=None;
        for (i,row) in rows.iter_mut().enumerate() {
            let cell=self.row(Grouping::Bare,Vec::new())?;if i==location.row{target=Some(cell.id);}
            let row=Arc::make_mut(row);row.origin=None;
            let NodeKind::GridRow{cells,separators,..}=&mut row.kind else{return Err(invalid_position())};
            cells.insert(at,cell);*separators=vec!["&".into();columns.len()-1];
        }
        self.update(location.grid,move|grid|{
            let NodeKind::Grid{rows:stored,columns:stored_columns,..}=&mut grid.kind else{return Err(invalid_position())};
            *stored=rows;*stored_columns=columns;Ok(())
        })?;
        self.selection=Selection::caret(Position::row(target.ok_or_else(invalid_position)?,0));Ok(())
    }
    pub(super) fn delete_grid_column(&mut self)->EngineResult<()> {
        let location=self.matrix_location()?;self.require_resizable_columns(location.grid)?;
        let (mut columns,mut rows)=self.matrix_parts(location.grid)?;
        if columns.len()==1 {return Err(EngineError::new(ErrorCode::InvalidInput,"A matrix must retain at least one column"));}
        columns.remove(location.column);let at=location.column.min(columns.len()-1);let mut target=None;
        for (i,row) in rows.iter_mut().enumerate() {
            let row=Arc::make_mut(row);row.origin=None;
            let NodeKind::GridRow{cells,separators,..}=&mut row.kind else{return Err(invalid_position())};
            let removed=cells.remove(location.column);let fallback=Position::row(cells[at].id,0);
            self.removed_matrix_node(&removed,fallback);if i==location.row{target=Some(fallback);}
            *separators=vec!["&".into();columns.len()-1];
        }
        self.update(location.grid,move|grid|{
            let NodeKind::Grid{rows:stored,columns:stored_columns,..}=&mut grid.kind else{return Err(invalid_position())};
            *stored=rows;*stored_columns=columns;Ok(())
        })?;
        self.selection=Selection::caret(target.ok_or_else(invalid_position)?);Ok(())
    }
    pub(super) fn move_grid_cell(&mut self,direction:Horizontal)->EngineResult<()> {
        let location=self.matrix_location()?;let (columns,rows)=self.matrix_parts(location.grid)?;
        let old=location.row*columns.len()+location.column;
        let target=match direction {Horizontal::Left=>old.saturating_sub(1),Horizontal::Right=>(old+1).min(rows.len()*columns.len()-1)};
        // Edge Tab is contained: it does not append a row or move host focus.
        let cell=&row_cells(&rows[target/columns.len()])?[target%columns.len()];
        self.selection=Selection::caret(Position::row(cell.id,0));Ok(())
    }
    pub(super) fn matrix_split_line(&mut self)->EngineResult<bool> {
        let Some(location)=nearest_matrix(&self.tree,&self.selection)? else{return Ok(false)};
        self.insert_matrix_row_at(location,true)?;Ok(true)
    }
    pub(super) fn matrix_delete_boundary(&self,direction:Horizontal)->EngineResult<bool> {
        let Some(location)=nearest_matrix(&self.tree,&self.selection)? else{return Ok(false)};
        let Some(Position::Row{node_id,offset,..})=self.selection.focus() else{return Ok(false)};
        if node_id!=location.cell{return Ok(false);}
        let entries=self.row_entries(node_id)?;
        Ok(match direction{Horizontal::Left=>offset==0,Horizontal::Right=>offset as usize==entries.len()})
    }
}
