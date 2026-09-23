//! Independent P08 review. These assertions exercise the public core and do
//! not replace keyboard, IME or two-frame browser acceptance. Implementation
//! files are owned by the concurrent contained-grid lane; keep failures visible.
use std::collections::BTreeSet;
use std::sync::Arc;
use serde_json::{json, Value};
use visualtex_engine_core::{
    grids::{GridColumnId, GridKind},
    model::{self, Grouping, IdAllocator, LatexMode, Node, NodeKind, RowEntry},
    parser,
    positions::{Position, TreeIndex},
    session::{EngineSession, NodeId},
    source::SourceText,
    ErrorCode, Runtime,
};

fn cell(id: u32) -> Arc<Node> { Node::empty_row(NodeId(id), Grouping::Bare) }
fn grid_row(id: u32, cell_id: u32, last: bool) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::GridRow {
        cells: vec![cell(cell_id)], separators: vec![],
        row_break: if last { String::new() } else { "\\\\".into() },
    } })
}
fn grid() -> Arc<Node> {
    Arc::new(Node { id: NodeId(2), origin: None, kind: NodeKind::Grid {
        environment: GridKind::Gathered, columns: vec![GridColumnId(3)], array_spec: None,
        before_begin_name: String::new(), before_end_name: String::new(),
        rows: vec![grid_row(4, 5, false), grid_row(6, 7, true)],
    } })
}
fn root_with(child: Arc<Node>) -> Arc<Node> {
    Arc::new(Node { id: NodeId(1), origin: None, kind: NodeKind::Row {
        grouping: Grouping::Bare,
        entries: vec![RowEntry { leading: String::new(), node: child }],
        trailing: String::new(),
    } })
}
fn parse(source: &str) -> parser::ParseResult {
    parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn identities(root: &Node) -> BTreeSet<u32> {
    let mut all = BTreeSet::new();
    root.visit(&mut |node| {
        assert!(all.insert(node.id.0), "Repeated live node {}", node.id.0);
        if let NodeKind::Grid { columns, .. } = &node.kind {
            for column in columns { assert!(all.insert(column.0), "Column identity collides with a node"); }
        }
    });
    all
}
fn semantics(value: Value) -> Value {
    match value {
        Value::Object(map) => Value::Object(map.into_iter()
            .filter(|(k, _)| !matches!(k.as_str(), "id" | "origin" | "columns" | "grouping" | "leading" | "trailing" | "beforeBeginName" | "beforeEndName"))
            .map(|(k, v)| (k, semantics(v))).collect()),
        Value::Array(values) => Value::Array(values.into_iter().map(semantics).collect()),
        value => value,
    }
}

#[test]
fn empty_cells_have_real_carets_but_column_metadata_is_not_a_caret() {
    let root = root_with(grid());
    let facts = model::validate_tree(&root).unwrap();
    assert_eq!(facts.grids, 1); assert_eq!(facts.grid_rows, 2); assert_eq!(facts.grid_columns, 1);
    assert_eq!(identities(&root).len(), facts.nodes + facts.grid_columns);
    let index = TreeIndex::new(&root).unwrap();
    index.validate_position(Position::row(NodeId(5), 0)).unwrap();
    index.validate_position(Position::row(NodeId(7), 0)).unwrap();
    assert!(index.validate_position(Position::row(NodeId(3), 0)).is_err());
    assert!(index.validate_position(Position::row(NodeId(4), 0)).is_err());
}

#[test]
fn duplicate_and_zero_column_identities_fail_instead_of_aliasing_cells() {
    for bad_column in [0, 1, 2, 4, 5, 6, 7] {
        let mut g = grid();
        let NodeKind::Grid { columns, .. } = &mut Arc::make_mut(&mut g).kind else { unreachable!() };
        columns[0] = GridColumnId(bad_column);
        assert_eq!(model::validate_tree(&root_with(g)).unwrap_err().code, ErrorCode::InvalidInput);
    }
}

#[test]
fn grid_rows_cannot_be_inserted_as_freestanding_math_atoms() {
    let orphan = root_with(grid_row(4, 5, true));
    assert!(model::validate_tree(&orphan).is_err(), "GridRow outside a Grid has no environment ownership");
}

#[test]
fn wrong_column_count_and_nonfinal_missing_break_are_rejected() {
    let mut g = grid();
    let NodeKind::Grid { rows, .. } = &mut Arc::make_mut(&mut g).kind else { unreachable!() };
    let NodeKind::GridRow { row_break, .. } = &mut Arc::make_mut(&mut rows[0]).kind else { unreachable!() };
    row_break.clear();
    assert!(model::validate_tree(&root_with(g)).is_err());
    let mut g = grid();
    let NodeKind::Grid { rows, .. } = &mut Arc::make_mut(&mut g).kind else { unreachable!() };
    let NodeKind::GridRow { cells, separators, .. } = &mut Arc::make_mut(&mut rows[0]).kind else { unreachable!() };
    cells.push(cell(8)); separators.push("&".into());
    assert!(model::validate_tree(&root_with(g)).is_err(), "gathered cannot hide a second column");
}

#[test]
fn column_allocation_is_monotonic_with_nodes_and_exhaustion_is_atomic() {
    let mut ids = IdAllocator::new(2);
    assert_eq!(GridColumnId::allocate(&mut ids).unwrap(), GridColumnId(2));
    assert_eq!(ids.allocate().unwrap(), NodeId(3));
    let mut ids = IdAllocator::new(u32::MAX);
    assert!(GridColumnId::allocate(&mut ids).is_err());
    assert_eq!(ids.next_unused(), u32::MAX);
}

#[test]
fn environment_newlines_comments_nested_groups_and_unknowns_are_never_lost() {
    let sources = [
        "\\begin{gathered}\nx+y=1\\\\\nx-y=0\n\\end{gathered}",
        "\\begin % before\r\n {gathered}a% hidden \\\\\n+b\\\\c\\end % end\n {gathered}",
        "\\begin{gathered}\\frac{a+b}{c}\\\\\\sqrt[3]{d}\\end{gathered}",
        "\\begin{gathered}\\text{A\\&B}\\\\\\future{x}\\end{gathered}",
        "\\begin{gathered}\\begin{gathered}a\\\\b\\end{gathered}\\\\c\\end{gathered}",
    ];
    for source in sources {
        let p = parse(source);
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
        let mut cursor = 0;
        for token in &p.tokens {
            assert_eq!(token.span.byte_start as usize, cursor);
            cursor = token.span.byte_end as usize;
        }
        assert_eq!(cursor, source.len());
        assert_eq!(identities(&p.tree).len(), p.facts.nodes + p.facts.grid_columns);
        if source.contains("\\future") { assert!(!p.supported && !p.exportable); }
        else {
            assert!(p.complete && p.supported, "{source}: {:?}", p.diagnostics);
            let canonical = model::latex(&p.tree, LatexMode::Canonical);
            let reopened = parse(&canonical);
            assert!(reopened.complete && reopened.supported, "{canonical}");
            assert_eq!(semantics(serde_json::to_value(&p.tree).unwrap()), semantics(serde_json::to_value(&reopened.tree).unwrap()));
        }
    }
}

#[test]
fn incomplete_and_mismatched_environment_input_remains_a_recoverable_draft() {
    for source in ["\\begin{gathered}x", "\\begin{gathered}x\\end{matrix}",
        "\\begin{gathered}\\frac{x}{\\end{gathered}", "\\begin{gathered}a&b\\end{gathered}"] {
        let p = parse(source);
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
        assert!(!p.exportable, "Invalid environment was made exportable: {source}");
    }
}

#[test]
fn complete_multiline_source_is_one_session_and_leaves_a_second_session_unchanged() {
    let mut runtime = Runtime::default();
    let send = |runtime: &mut Runtime, command: Value| -> Value {
        serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
    };
    assert_eq!(send(&mut runtime, json!({"op":"create","documentId":"frame-A","source":"a+b"}))["status"], "ok");
    assert_eq!(send(&mut runtime, json!({"op":"create","documentId":"frame-B","source":"z^2"}))["status"], "ok");
    let b_before = send(&mut runtime, json!({"op":"inspectEditor","handle":2}));
    let source = "\\begin{gathered}a+b\\\\c+d\\end{gathered}";
    let a = send(&mut runtime, json!({"op":"replaceSource","handle":1,"expectedRevision":0,"source":source}));
    assert_eq!(a["status"], "ok");
    assert_eq!(a["result"]["snapshot"]["documentId"], "frame-A");
    assert_eq!(a["result"]["snapshot"]["analysis"]["facts"]["grids"], 1);
    assert_eq!(runtime.active_sessions(), 2);
    assert_eq!(send(&mut runtime, json!({"op":"inspectEditor","handle":2})), b_before);
    let undo = send(&mut runtime, json!({"op":"editorAction","handle":1,"expectedRevision":1,"eventSequence":1,"action":{"kind":"undo"}}));
    assert_eq!(undo["status"], "ok");
    assert_eq!(undo["result"]["result"]["editor"]["document"]["source"], "a+b");
    assert_eq!(send(&mut runtime, json!({"op":"inspectEditor","handle":2})), b_before);
}

#[test]
fn reloaded_multiline_snapshot_retains_single_document_identity() {
    let source = "\\begin{gathered}a\\\\b\\end{gathered}";
    let s = EngineSession::from_source("one-formula".into(), source.into()).unwrap();
    let snapshot = s.document_snapshot();
    assert_eq!(snapshot.document_id, "one-formula");
    assert_eq!(snapshot.analysis.facts.grids, 1);
    assert_eq!(snapshot.analysis.facts.grid_rows, 2);
    assert_eq!(snapshot.source, source);
}
