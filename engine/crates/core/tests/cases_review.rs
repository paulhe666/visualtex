//! P08 cases: one environment owner, real brace geometry, identified rows/cells.
//! Native tests are not browser, native IME or installed Office acceptance.
use std::{collections::BTreeSet, path::PathBuf};
use serde_json::{json, Value};
use visualtex_engine_core::{editing::{EditorAction, EditorActionResult}, fonts::FontRegistry,
    grids::GridKind, layout::{LayoutEngine,LayoutOptions,MathStyle}, model::{self,Node,NodeKind,LatexMode},
    positions::{Horizontal,Position,Selection,TreeIndex}, session::EngineSession,
    transactions::{EditCommand,InputMode}};

struct Owned { s:EngineSession, seq:u32 }
impl Owned {
    fn new(source:&str)->Self {Self{s:EngineSession::from_source("cases-review".into(),source.into()).unwrap(),seq:0}}
    fn snapshot(&self)->Value {serde_json::to_value(self.s.editor_snapshot().unwrap()).unwrap()}
    fn action(&mut self,a:EditorAction)->EditorActionResult {
        self.seq+=1;let old=self.s.document_snapshot();
        let result=self.s.apply_action(self.s.revision(),self.seq,a).unwrap();
        let tree=&result.editor.document.analysis.tree;let index=TreeIndex::new(tree).unwrap();
        index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map)=&result.position_map {for p in TreeIndex::new(&old.analysis.tree).unwrap().carets {
            index.validate_position(map.apply_position(p,&index).unwrap()).unwrap();
        }}
        let fresh=Owned::new(&result.editor.document.source).s.document_snapshot();
        assert_eq!(semantics(serde_json::to_value(tree).unwrap()),semantics(serde_json::to_value(&fresh.analysis.tree).unwrap()));
        result
    }
    fn edit(&mut self,c:EditCommand)->EditorActionResult {self.action(EditorAction::Edit{command:c,group_id:None})}
    fn select(&mut self,p:Position) {self.action(EditorAction::SetSelection{selection:Selection::caret(p)});}
    fn reject(&mut self,c:EditCommand) {let before=self.snapshot();self.seq+=1;
        assert!(self.s.apply_action(self.s.revision(),self.seq,EditorAction::Edit{command:c,group_id:None}).is_err());
        assert_eq!(self.snapshot(),before);
    }
}
fn semantics(v:Value)->Value {match v {
    Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"leading"|"trailing"|"grouping")&&!k.starts_with("before"))
        .map(|(k,v)| {let v=if k=="columns" {json!(v.as_array().unwrap().len())}else{semantics(v)};(k,v)}).collect()),
    Value::Array(a)=>Value::Array(a.into_iter().map(semantics).collect()),v=>v}}
fn content(v:&Value)->Value {json!({"source":v["document"]["source"],"tree":v["document"]["analysis"]["tree"],"selection":v["selection"],"owner":v["document"]["documentId"]})}
fn grid(n:&Node)->&Node {if matches!(n.kind,NodeKind::Grid{..}){return n;} n.children().into_iter().find_map(|c|{
    let mut found=false;c.visit(&mut |n|found|=matches!(n.kind,NodeKind::Grid{..}));found.then(||grid(c))}).expect("real grid")}
fn cell(g:&Node,r:usize,c:usize)->&Node {let NodeKind::Grid{rows,..}=&g.kind else{panic!()};
    let NodeKind::GridRow{cells,..}=&rows[r].kind else{panic!()}; &cells[c]}
fn ids(n:&Node)->BTreeSet<u32> {let mut ids=BTreeSet::new();n.visit(&mut |n| {assert!(ids.insert(n.id.0));
    if let NodeKind::Grid{columns,..}=&n.kind {for c in columns {assert!(ids.insert(c.0));}}});ids}
fn font()->(FontRegistry,LayoutOptions) {let p=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(p).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}
const SOURCE:&str=r"\begin{cases}x^2&x\ge0\\-x&x<0\end{cases}";

#[test]
fn authored_cases_keep_two_real_columns_and_lossless_roundtrip() {
    for source in [SOURCE,r"\begin{cases}&\\&\end{cases}",r"\begin{cases}\frac{x}{y}&\text{if }x>0\\\sqrt{x}&\text{otherwise}\end{cases}",
        "\\begin %header\n {cases}a% hidden & \\\\\n&b\\\\c&d\\end %footer\n {cases}",
        r"\begin{cases}\begin{matrix}a&b\\c&d\end{matrix}&x>0\\\begin{gathered}p\\q\end{gathered}&x<0\end{cases}"] {
        let s=Owned::new(source);let d=s.s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported,"{:?}",d.analysis.diagnostics);
        assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
        let g=grid(&d.analysis.tree);let NodeKind::Grid{environment,columns,..}=&g.kind else{panic!()};
        assert_eq!(*environment,GridKind::Cases);assert_eq!(columns.len(),2);
        assert_eq!(ids(&d.analysis.tree).len(),d.analysis.facts.nodes+d.analysis.facts.grid_columns);
        let canonical=model::latex(&d.analysis.tree,LatexMode::Canonical);assert!(Owned::new(&canonical).s.document_snapshot().analysis.supported);
    }
}
#[test]
fn cases_insertion_preserves_selected_children_and_single_history_item() {
    let mut t=Owned::new("xy");t.action(EditorAction::SetSelection{selection:Selection::Range{
        anchor:Position::row(t.s.document_snapshot().analysis.tree.id,0),focus:Position::row(t.s.document_snapshot().analysis.tree.id,2)}});
    let before=t.snapshot();let old=t.s.document_snapshot();let NodeKind::Row{entries,..}=&old.analysis.tree.kind else{panic!()};
    let expected:Vec<_>=entries.iter().map(|e|e.node.id).collect();let after=t.edit(EditCommand::InsertCases{rows:2}).editor;
    let NodeKind::Row{entries,..}=&cell(grid(&after.document.analysis.tree),0,0).kind else{panic!()};
    assert_eq!(entries.iter().map(|e|e.node.id).collect::<Vec<_>>(),expected);
    assert_eq!(after.undo_depth,before["undoDepth"].as_u64().unwrap() as usize+1);let saved=t.snapshot();
    t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));
    t.action(EditorAction::Redo);assert_eq!(content(&t.snapshot()),content(&saved));
}
#[test]
fn cases_row_commands_keep_column_schema_and_exact_history() {
    for after in [false,true] {let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
        t.select(Position::row(cell(g,0,1).id,1));let before=t.snapshot();let old_ids=ids(g);
        let added=t.edit(EditCommand::InsertGridRow{after}).editor;let next=grid(&added.document.analysis.tree);
        assert_eq!(added.document.analysis.facts.grid_rows,3);assert!(old_ids.is_subset(&ids(next)));
        assert_eq!(added.selection.focus().unwrap().node_id(),cell(next,usize::from(after),1).id);
        let saved=t.snapshot();t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));
        t.action(EditorAction::Redo);assert_eq!(content(&t.snapshot()),content(&saved));
        t.edit(EditCommand::DeleteGridRow);assert_eq!(t.s.source(),SOURCE);
        t.reject(EditCommand::InsertGridColumn{after:true});t.reject(EditCommand::DeleteGridColumn);
    }
}
#[test]
fn cases_nearest_environment_insertion_does_not_cut_nested_fraction() {
    let mut t=Owned::new(r"\begin{matrix}a&\begin{cases}\frac{xy}{z}&q\\c&d\end{cases}\\e&f\end{matrix}");
    let old=t.s.document_snapshot();let outer=grid(&old.analysis.tree);let inner=grid(cell(outer,0,1));
    let NodeKind::Row{entries,..}=&cell(inner,0,0).kind else{panic!()};let NodeKind::Fraction{numerator,..}=&entries[0].node.kind else{panic!()};
    let fraction_id=entries[0].node.id;t.select(Position::row(numerator.id,1));let before=t.snapshot();
    let d=t.edit(EditCommand::SplitLine).editor.document;let outer_next=grid(&d.analysis.tree);let inner_next=grid(cell(outer_next,0,1));
    assert_eq!(d.analysis.facts.grids,2);assert_eq!(outer_next.id,outer.id);assert_eq!(inner_next.id,inner.id);
    let NodeKind::Grid{rows,..}=&outer_next.kind else{panic!()};assert_eq!(rows.len(),2);
    let NodeKind::Grid{rows,..}=&inner_next.kind else{panic!()};assert_eq!(rows.len(),3);
    assert!(ids(inner_next).contains(&fraction_id.0));assert_eq!(model::latex(inner_next.find(fraction_id).unwrap(),LatexMode::Lossless),r"\frac{xy}{z}");
    t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));
}
#[test]
fn raw_cross_case_ranges_cannot_delete_or_copy_the_whole_cases_object() {
    let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
    t.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(cell(g,0,0).id,0),focus:Position::row(cell(g,1,1).id,1)}});
    t.reject(EditCommand::DeleteBackward);assert!(t.s.copy_selection(t.s.revision()).is_err());
}
#[test]
fn cases_rectangular_clipboard_preserves_cells_and_does_not_copy_the_brace() {
    let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
    let selection=Selection::GridCells{grid_id:g.id,anchor_cell:cell(g,0,0).id,focus_cell:cell(g,1,1).id};
    t.action(EditorAction::SetSelection{selection});let before=t.snapshot();let fragment=t.s.copy_selection(t.s.revision()).unwrap();
    assert!(fragment.latex.as_ref().unwrap().starts_with(r"\begin{matrix}"));assert_eq!(t.snapshot(),before);
    t.edit(EditCommand::ClearGridCells);assert_eq!(t.s.document_snapshot().analysis.facts.grid_rows,2);
    t.edit(EditCommand::PasteGridCells{latex:fragment.latex.unwrap()});let next=t.s.document_snapshot();
    assert_eq!(grid(&next.analysis.tree).id,g.id);for r in 0..2 {for c in 0..2 {assert_eq!(cell(grid(&next.analysis.tree),r,c).id,cell(g,r,c).id);}}
    assert_eq!(semantics(serde_json::to_value(g).unwrap()),semantics(serde_json::to_value(grid(&next.analysis.tree)).unwrap()));
}
#[test]
fn cases_cell_navigation_and_edge_delete_remain_internal() {
    let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
    t.select(Position::row(cell(g,0,0).id,0));let before=t.snapshot();t.edit(EditCommand::DeleteBackward);assert_eq!(t.snapshot(),before);
    t.edit(EditCommand::MoveGridCell{direction:Horizontal::Right});assert_eq!(t.s.editor_snapshot().unwrap().selection.focus().unwrap().node_id(),cell(g,0,1).id);
    t.edit(EditCommand::MoveGridCell{direction:Horizontal::Right});assert_eq!(t.s.editor_snapshot().unwrap().selection.focus().unwrap().node_id(),cell(g,1,0).id);
    assert_eq!(t.s.revision().0,0);assert_eq!(t.s.source(),SOURCE);
}
#[test]
fn cases_brace_uses_real_font_glyphs_and_left_columns_share_anchors() {
    let (fonts,mut options)=font();let t=Owned::new(r"\begin{cases}abc&x>0\\d&x<0\end{cases}");
    let mut engine=LayoutEngine::default();let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
    for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript] {
        options.style=style;options.font_size_px=16.0;let a=engine.layout(&t.s,&fonts,&options).unwrap();a.validate().unwrap();
        let brace:Vec<_>=a.glyphs.iter().filter(|glyph|glyph.node_id==g.id).collect();assert!(!brace.is_empty());
        for c in 0..2 {let x=a.caret(Position::row(cell(g,0,c).id,0)).unwrap().x;
            assert!((a.caret(Position::row(cell(g,1,c).id,0)).unwrap().x-x).abs()<1e-9);
            assert!((a.grid_columns[c].x-x).abs()<1e-9);
        }
        assert!(a.grid_columns[0].x>0.0);let svg=a.svg().unwrap();assert_eq!(svg.matches("<svg").count(),1);
        for bad in ["<text","foreignObject","NaN","Infinity"] {assert!(!svg.contains(bad));}
        options.font_size_px=32.0;let b=engine.layout(&t.s,&fonts,&options).unwrap();assert_eq!(b.width,a.width*2.0);assert_eq!(b.height,a.height*2.0);
        assert_eq!(serde_json::to_value(b.grid_columns).unwrap(),serde_json::to_value(a.grid_columns).unwrap());
    }
}
#[test]
fn cases_invalid_shape_and_row_budget_are_atomic() {
    for source in [r"\begin{cases}x&y&z\end{cases}",r"\begin{cases}x\\a&b\end{cases}",r"\begin{cases}x&y\\[2pt]a&b\end{cases}",r"\begin{cases}x&y"] {
        let t=Owned::new(source);assert_eq!(t.s.source(),source);assert!(!t.s.document_snapshot().analysis.exportable);
    }
    let mut t=Owned::new(r"\begin{cases}x&y\end{cases}");let d=t.s.document_snapshot();t.select(Position::row(cell(grid(&d.analysis.tree),0,0).id,0));
    t.reject(EditCommand::DeleteGridRow);t.reject(EditCommand::InsertCases{rows:0});t.reject(EditCommand::InsertCases{rows:257});
}
#[test]
fn repeated_cases_row_edits_keep_one_identity_and_fresh_parser_semantics() {
    let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);let grid_id=g.id;
    t.select(Position::row(cell(g,1,0).id,0));
    for i in 0..32 {t.edit(EditCommand::SplitLine);t.edit(EditCommand::InsertText{text:i.to_string(),mode:InputMode::Math});
        assert_eq!(grid(&t.s.document_snapshot().analysis.tree).id,grid_id);assert_eq!(t.s.document_snapshot().analysis.facts.grids,1);
    }
}
