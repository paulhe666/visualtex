//! Rectangles select identified cells in one grid, never source offsets or hosts.
use serde::{Deserialize, Serialize};
use crate::{error::{EngineResult, EngineError, ErrorCode}, model::NodeKind,
    positions::{Position, Selection, TreeIndex, invalid_position}, session::NodeId};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all="camelCase")]
pub enum GridDirection { Left, Right, Up, Down }

#[derive(Debug, Clone, Copy)]
pub struct CellLocation { pub grid:NodeId, pub cell:NodeId, pub row:usize, pub column:usize }

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all="camelCase")]
pub struct GridRectangle {
    pub grid_id:NodeId, pub row_start:usize, pub row_end:usize,
    pub column_start:usize, pub column_end:usize,
}
impl GridRectangle {
    pub fn rows(&self)->usize { self.row_end-self.row_start+1 }
    pub fn columns(&self)->usize { self.column_end-self.column_start+1 }
}
impl TreeIndex<'_> {
    /// Return the nearest actual environment. Do not walk past a nested grid.
    pub fn grid_cell(&self,mut id:NodeId)->Option<CellLocation> {
        while let Some(parent)=self.nodes.get(&id)?.parent {
            if let NodeKind::GridRow{cells,..}=&self.nodes.get(&parent)?.node.kind {
                let grid=self.nodes.get(&parent)?.parent?;
                let NodeKind::Grid{rows,..}=&self.nodes.get(&grid)?.node.kind else{return None};
                return Some(CellLocation{grid,cell:id,row:rows.iter().position(|r|r.id==parent)?,
                    column:cells.iter().position(|c|c.id==id)?});
            }
            id=parent;
        }
        None
    }
    pub fn grid_rectangle(&self,grid:NodeId,anchor:NodeId,focus:NodeId)->EngineResult<GridRectangle> {
        let a=self.grid_cell(anchor).ok_or_else(invalid_position)?;
        let b=self.grid_cell(focus).ok_or_else(invalid_position)?;
        // Endpoints must be the cell nodes themselves, not descendants or a
        // same-number coordinate borrowed from an unrelated nested matrix.
        if a.grid!=grid || b.grid!=grid || a.cell!=anchor || b.cell!=focus {
            return Err(invalid_position());
        }
        Ok(GridRectangle{grid_id:grid,row_start:a.row.min(b.row),row_end:a.row.max(b.row),
            column_start:a.column.min(b.column),column_end:a.column.max(b.column)})
    }
    pub fn cell_at(&self,grid:NodeId,row:usize,column:usize)->EngineResult<NodeId> {
        let NodeKind::Grid{rows,..}=&self.node(grid)?.kind else{return Err(invalid_position())};
        let NodeKind::GridRow{cells,..}=&rows.get(row).ok_or_else(invalid_position)?.kind else{return Err(invalid_position())};
        Ok(cells.get(column).ok_or_else(invalid_position)?.id)
    }
    /// A pointer span entirely within one cell stays a mathematical range.
    /// Different cells of the same grid select whole cells, not the outer atom.
    pub fn pointer_selection(&self,anchor:Position,focus:Position)->EngineResult<Selection> {
        self.validate_position(anchor)?;self.validate_position(focus)?;
        if let (Some(a),Some(b))=(self.grid_cell(anchor.node_id()),self.grid_cell(focus.node_id())) {
            if a.grid==b.grid && a.cell!=b.cell {
                return Ok(Selection::GridCells{grid_id:a.grid,anchor_cell:a.cell,focus_cell:b.cell});
            }
        }
        Ok(Selection::Range{anchor,focus})
    }
    pub fn extend_grid_selection(&self,selection:&Selection,direction:GridDirection)->EngineResult<Selection> {
        self.validate_selection(selection)?;
        let (grid,anchor,focus)=match selection {
            Selection::GridCells{grid_id,anchor_cell,focus_cell}=>(*grid_id,*anchor_cell,*focus_cell),
            Selection::Range{anchor,focus} if anchor.same_location(*focus)=>{
                let cell=self.grid_cell(focus.node_id()).ok_or_else(invalid_position)?;
                (cell.grid,cell.cell,cell.cell)
            }
            _=>return Err(EngineError::new(ErrorCode::InvalidPosition,"Cell extension requires a caret or a cell rectangle")),
        };
        let cell=self.grid_cell(focus).ok_or_else(invalid_position)?;
        let NodeKind::Grid{rows,columns,..}=&self.node(grid)?.kind else{return Err(invalid_position())};
        let (r,c)=match direction {
            GridDirection::Left=>(cell.row,cell.column.saturating_sub(1)),
            GridDirection::Right=>(cell.row,(cell.column+1).min(columns.len()-1)),
            GridDirection::Up=>(cell.row.saturating_sub(1),cell.column),
            GridDirection::Down=>((cell.row+1).min(rows.len()-1),cell.column),
        };
        Ok(Selection::GridCells{grid_id:grid,anchor_cell:anchor,focus_cell:self.cell_at(grid,r,c)?})
    }
}
