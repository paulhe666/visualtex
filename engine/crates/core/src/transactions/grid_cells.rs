//! Cell-content replacement is atomic and preserves the destination grid, row,
//! column and cell identities. Clipboard source creates fresh descendant IDs.
use super::*;
use crate::grid_selection::GridRectangle;

impl Builder<'_> {
    fn cell_rectangle(&self)->EngineResult<GridRectangle> {
        let Selection::GridCells{grid_id,anchor_cell,focus_cell}=self.selection else {
            return Err(EngineError::new(ErrorCode::InvalidPosition,"Select a cell rectangle first"));
        };
        TreeIndex::new(&self.tree)?.grid_rectangle(grid_id,anchor_cell,focus_cell)
    }
    pub(super) fn clear_grid_cells(&mut self)->EngineResult<()> {
        let rect=self.cell_rectangle()?;
        self.replace_grid_cells(&rect,None)
    }
    pub(super) fn paste_grid_cells(&mut self,latex:&str,source_id:u32)->EngineResult<()> {
        let root_id=self.ids.allocate()?;
        let parsed=parser::parse_with_symbols(SourceText::new(source_id,latex)?,root_id,self.ids,Arc::clone(&self.symbols))?;
        if !parsed.complete {return Err(EngineError::new(ErrorCode::IncompleteSource,"Cell clipboard contains unfinished source"));}
        if !parsed.supported {return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Cell clipboard contains unsupported source"));}
        let NodeKind::Row{entries,trailing,..}=&parsed.tree.kind else{return Err(invalid_position())};
        let source:Vec<Vec<Arc<Node>>>;
        if entries.len()==1 && matches!(entries[0].node.kind,NodeKind::Grid{..}) {
            // Outer comments cannot be silently discarded by extracting a grid.
            if !entries[0].leading.trim().is_empty() || !trailing.trim().is_empty() {
                return Err(EngineError::new(ErrorCode::InvalidInput,"Put clipboard comments inside cells"));
            }
            let NodeKind::Grid{environment:crate::grids::GridKind::Matrix,rows,..}=&entries[0].node.kind else {
                return Err(EngineError::new(ErrorCode::InvalidInput,"Rectangular clipboard requires matrix cell data"));
            };
            source=rows.iter().map(|row|match &row.kind {
                NodeKind::GridRow{cells,..}=>Ok(cells.clone()),_=>Err(invalid_position())
            }).collect::<EngineResult<_>>()?;
        } else { source=vec![vec![Arc::clone(&parsed.tree)]]; }
        let nr=source.len();let nc=source[0].len();
        let index=TreeIndex::new(&self.tree)?;
        let rect=if matches!(self.selection,Selection::GridCells{..}) {
            let rect=self.cell_rectangle()?;
            if rect.rows()!=nr || rect.columns()!=nc {
                return Err(EngineError::new(ErrorCode::InvalidInput,"Clipboard dimensions must exactly match the selected rectangle"));
            }
            rect
        } else {
            if !self.selection.is_collapsed(){return Err(invalid_position());}
            let location=index.grid_cell(self.selection.focus().ok_or_else(invalid_position)?.node_id()).ok_or_else(invalid_position)?;
            let NodeKind::Grid{rows,columns,..}=&index.node(location.grid)?.kind else{return Err(invalid_position())};
            if location.row+nr>rows.len() || location.column+nc>columns.len() {
                return Err(EngineError::new(ErrorCode::InvalidInput,"Clipboard rectangle exceeds the current grid"));
            }
            GridRectangle{grid_id:location.grid,row_start:location.row,row_end:location.row+nr-1,
                column_start:location.column,column_end:location.column+nc-1}
        };
        let anchor=index.cell_at(rect.grid_id,rect.row_start,rect.column_start)?;
        let focus=index.cell_at(rect.grid_id,rect.row_end,rect.column_end)?;
        self.replace_grid_cells(&rect,Some(&source))?;
        self.selection=Selection::GridCells{grid_id:rect.grid_id,anchor_cell:anchor,focus_cell:focus};
        Ok(())
    }
    fn replace_grid_cells(&mut self,rect:&GridRectangle,source:Option<&Vec<Vec<Arc<Node>>>>)->EngineResult<()> {
        let NodeKind::Grid{rows,..}=&self.tree.find(rect.grid_id).ok_or_else(invalid_position)?.kind else{return Err(invalid_position())};
        let mut rows=rows.clone();let mut changed=false;
        for r in rect.row_start..=rect.row_end {
            let row=Arc::make_mut(&mut rows[r]);
            let NodeKind::GridRow{cells,..}=&mut row.kind else{return Err(invalid_position())};
            for c in rect.column_start..=rect.column_end {
                let old=&cells[c];
                let NodeKind::Row{entries:old_entries,trailing:old_trailing,..}=&old.kind else{return Err(invalid_position())};
                let (entries,trailing)=if let Some(source)=source {
                    let NodeKind::Row{entries,trailing,..}=&source[r-rect.row_start][c-rect.column_start].kind else{return Err(invalid_position())};
                    (entries.clone(),trailing.clone())
                }else{(Vec::new(),String::new())};
                if source.is_none() && old_entries.is_empty() && old_trailing.is_empty(){continue;}
                let id=old.id;let fallback=Position::row(id,0);
                for entry in old_entries {entry.node.visit(&mut |n|self.overrides.push(PositionStep::Removed{node_id:n.id,fallback}));}
                // PositionMap::between derives the retained cell's splice map;
                // descendant fallbacks above cover removed inner structures.
                cells[c]=Arc::new(Node{id,origin:None,kind:NodeKind::Row{grouping:Grouping::Bare,entries,trailing}});
                changed=true;row.origin=None;
            }
        }
        if changed {self.update(rect.grid_id,move|grid|{
            let NodeKind::Grid{rows:stored,..}=&mut grid.kind else{return Err(invalid_position())};
            *stored=rows;Ok(())
        })?;}
        Ok(())
    }
}
