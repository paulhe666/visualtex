//! Validate contained-grid alignment metadata independently of glyph drawing.
//! Coordinates are scene em units. Hidden phantom grids may lie outside the
//! visible export frame, but every column remains inside its own grid region.
use std::collections::{BTreeMap, BTreeSet};
use crate::{
    error::{EngineError, EngineResult, ErrorCode},
    grids::MAX_GRID_COLUMNS,
    model::MAX_NODES,
    scene::{GridColumnAnchor, LayoutSnapshot},
    session::NodeId,
};

fn invalid(message: &str) -> EngineError { EngineError::new(ErrorCode::InvalidInput, message) }

pub(crate) fn validate(snapshot: &LayoutSnapshot) -> EngineResult<()> {
    if snapshot.grid_columns.is_empty() && !snapshot.regions.iter().any(|region| region.role == "grid") {
        return Ok(());
    }
    if snapshot.grid_columns.len() > MAX_NODES || snapshot.regions.len() > MAX_NODES {
        return Err(EngineError::new(ErrorCode::ResourceLimit, "Contained-grid scene metadata budget exhausted"));
    }
    let mut nodes = BTreeSet::new();
    let mut owners = BTreeMap::new();
    for region in &snapshot.regions {
        if region.node_id.0 == 0 || !nodes.insert(region.node_id) {
            return Err(invalid("Grid scene contains duplicate or zero mathematical region identities"));
        }
        if region.role == "grid" {
            if !region.logical.valid() { return Err(invalid("Grid owner has invalid logical bounds")); }
            owners.insert(region.node_id, region);
        }
    }
    let mut columns = BTreeSet::new();
    let mut groups: BTreeMap<NodeId, BTreeMap<usize, &GridColumnAnchor>> = BTreeMap::new();
    for column in &snapshot.grid_columns {
        if column.column_id.0 == 0 || nodes.contains(&NodeId(column.column_id.0)) || !columns.insert(column.column_id) {
            return Err(invalid("Grid column identity is duplicate, zero or aliases a mathematical node"));
        }
        let owner = owners.get(&column.grid_id)
            .ok_or_else(|| invalid("Grid column has no matching mathematical grid owner"))?;
        if column.index >= MAX_GRID_COLUMNS || !column.x.is_finite() || !column.width.is_finite()
            || column.width < 0.0 || column.x.abs() > 1_000_000.0 || column.width > 1_000_000.0 {
            return Err(invalid("Invalid contained-grid column coordinate, width or index"));
        }
        let right = column.x + column.width;
        let epsilon = 1e-9 * owner.logical.width().max(column.x.abs()).max(right.abs()).max(1.0);
        if !right.is_finite() || column.x < owner.logical.x_min - epsilon || right > owner.logical.x_max + epsilon {
            return Err(invalid("Column anchor lies outside its owning grid's logical bounds"));
        }
        if groups.entry(column.grid_id).or_default().insert(column.index, column).is_some() {
            return Err(invalid("Duplicate column index within one mathematical grid"));
        }
    }
    if groups.len() != owners.len() {
        return Err(invalid("A mathematical grid has no column anchors"));
    }
    for group in groups.values() {
        let mut previous_right: Option<f64> = None;
        for (expected_index, (index, column)) in group.iter().enumerate() {
            if *index != expected_index { return Err(invalid("Grid column indices are not contiguous from zero")); }
            if let Some(right) = previous_right {
                let epsilon = 1e-9 * right.abs().max(column.x.abs()).max(1.0);
                if column.x < right - epsilon { return Err(invalid("Grid column anchors overlap or are out of order")); }
            }
            previous_right = Some(column.x + column.width);
        }
    }
    Ok(())
}
