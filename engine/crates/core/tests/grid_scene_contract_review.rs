//! Independent P08 scene/ownership checks. No browser DOM or Office geometry is
//! used as the oracle. Fonts are existing, read-only local test resources.
use std::path::PathBuf;
use visualtex_engine_core::{
    fonts::FontRegistry,
    grids::GridColumnId,
    layout::{LayoutEngine, LayoutOptions},
    model::{Node, NodeKind},
    positions::{Position, TreeIndex},
    scene::LayoutSnapshot,
    session::{EngineSession, NodeId},
};

fn fixture(source: &str) -> (EngineSession, FontRegistry, LayoutOptions) {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts = FontRegistry::default();
    let id = fonts.register(&std::fs::read(path).unwrap(), 0).unwrap().id;
    let session = EngineSession::from_source("contained-scene-review".into(), source.into()).unwrap();
    (session, fonts, LayoutOptions::new(id))
}
fn first_grid(root: &Node) -> &Node {
    if matches!(root.kind, NodeKind::Grid { .. }) { return root; }
    for child in root.children() {
        let mut found = false;
        child.visit(&mut |n| { found |= matches!(n.kind, NodeKind::Grid { .. }); });
        if found { return first_grid(child); }
    }
    panic!("Expected a real Grid, not an opaque source string");
}
fn visible_scene() -> (EngineSession, LayoutSnapshot) {
    let (session, fonts, options) = fixture(r"\begin{gathered}x+y=1\\\frac{a}{b}=2\end{gathered}");
    let scene = LayoutEngine::default().layout(&session, &fonts, &options).unwrap();
    (session, scene)
}

#[test]
fn gathered_has_one_owner_and_separate_real_row_baselines() {
    let (session, scene) = visible_scene();
    let snapshot = session.document_snapshot();
    let grid = first_grid(&snapshot.analysis.tree);
    let NodeKind::Grid { columns, rows, .. } = &grid.kind else { unreachable!() };
    assert_eq!(rows.len(), 2); assert_eq!(columns.len(), 1);
    assert_eq!(scene.grid_columns.len(), 1);
    let anchor = &scene.grid_columns[0];
    assert_eq!(anchor.grid_id, grid.id); assert_eq!(anchor.column_id, columns[0]); assert_eq!(anchor.index, 0);
    let cell_ids: Vec<_> = rows.iter().map(|row| match &row.kind {
        NodeKind::GridRow { cells, .. } => cells[0].id,
        _ => panic!("Grid did not contain GridRows"),
    }).collect();
    let a = scene.caret(Position::row(cell_ids[0], 0)).unwrap();
    let b = scene.caret(Position::row(cell_ids[1], 0)).unwrap();
    assert!(b.baseline > a.baseline, "Contained rows collapsed onto the same baseline");
    assert!(scene.regions.iter().any(|r| r.node_id == grid.id));
    let index = TreeIndex::new(&snapshot.analysis.tree).unwrap();
    for caret in &scene.carets { index.validate_position(caret.position).unwrap(); }
    assert_eq!(scene.document_id, snapshot.document_id);
    assert_eq!(scene.source_id, snapshot.source_id);
    let svg = scene.svg().unwrap();
    assert_eq!(svg.matches("<svg").count(), 1);
    for invalid in ["<text", "<foreignObject", "NaN", "Infinity"] { assert!(!svg.contains(invalid)); }
}

#[test]
fn empty_grid_cells_remain_real_hit_test_targets() {
    let (session, fonts, options) = fixture(r"\begin{gathered}a\\\end{gathered}");
    let scene = LayoutEngine::default().layout(&session, &fonts, &options).unwrap();
    let snapshot = session.document_snapshot();
    let NodeKind::Grid { rows, .. } = &first_grid(&snapshot.analysis.tree).kind else { unreachable!() };
    let NodeKind::GridRow { cells, .. } = &rows.last().unwrap().kind else { unreachable!() };
    let position = Position::row(cells[0].id, 0);
    let stop = scene.caret(position).unwrap();
    let hit = scene.hit_test(stop.x, stop.baseline).unwrap();
    assert!(hit.same_location(position), "Empty internal row could not receive the caret");
}

#[test]
fn grid_scene_natural_size_and_column_units_follow_font_size_only_once() {
    let (session, fonts, mut options) = fixture(r"\begin{gathered}a\\b+c\end{gathered}");
    let mut engine = LayoutEngine::default();
    let first = engine.layout(&session, &fonts, &options).unwrap();
    options.font_size_px *= 2.0;
    let second = engine.layout(&session, &fonts, &options).unwrap();
    assert_eq!(second.width, first.width * 2.0);
    assert_eq!(second.height, first.height * 2.0);
    assert_eq!(second.baseline, first.baseline * 2.0);
    assert_eq!(second.width_em, first.width_em);
    assert_eq!(second.height_em, first.height_em);
    assert_eq!(serde_json::to_value(&first.grid_columns).unwrap(), serde_json::to_value(&second.grid_columns).unwrap());
    assert_eq!(second.rebuilt_fragments, 0);
}

#[test]
fn grid_column_anchors_reject_nan_infinity_and_negative_width() {
    let (_, scene) = visible_scene();
    assert!(!scene.grid_columns.is_empty());
    for bad in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let mut candidate = scene.clone(); candidate.grid_columns[0].x = bad;
        assert!(candidate.validate().is_err(), "Nonfinite column x accepted");
        assert!(candidate.svg().is_err(), "Nonfinite column x exported");
        let mut candidate = scene.clone(); candidate.grid_columns[0].width = bad;
        assert!(candidate.validate().is_err(), "Nonfinite column width accepted");
    }
    let mut candidate = scene.clone(); candidate.grid_columns[0].width = -1.0;
    assert!(candidate.validate().is_err(), "Negative grid column width accepted");
}

#[test]
fn a_grid_column_cannot_borrow_an_unrelated_region_or_duplicate_ownership() {
    let (session, scene) = visible_scene();
    let mut candidate = scene.clone(); candidate.grid_columns[0].grid_id = NodeId(0);
    assert!(candidate.validate().is_err(), "Missing grid owner accepted");
    let mut candidate = scene.clone(); candidate.grid_columns[0].grid_id = session.document_snapshot().analysis.tree.id;
    assert!(candidate.validate().is_err(), "Ordinary root row was accepted as an alignment owner");
    let mut candidate = scene.clone(); candidate.grid_columns[0].column_id = GridColumnId(0);
    assert!(candidate.validate().is_err(), "Zero column identity accepted");
    let mut candidate = scene.clone(); candidate.grid_columns.push(candidate.grid_columns[0].clone());
    assert!(candidate.validate().is_err(), "Duplicate grid/column anchor accepted");
    let mut candidate = scene.clone(); candidate.grid_columns[0].index = 8_192;
    assert!(candidate.validate().is_err(), "Non-contiguous or overbudget column index accepted");
}

#[test]
fn a_column_anchor_outside_its_grid_is_not_hidden_by_valid_glyph_bounds() {
    let (_, scene) = visible_scene();
    let mut candidate = scene.clone(); candidate.grid_columns[0].x = scene.width_em + 100.0;
    assert!(candidate.validate().is_err(), "Column anchor outside the grid/frame accepted");
    let mut candidate = scene.clone(); candidate.grid_columns[0].width = scene.width_em + 100.0;
    assert!(candidate.validate().is_err(), "Column width outside the grid/frame accepted");
}
