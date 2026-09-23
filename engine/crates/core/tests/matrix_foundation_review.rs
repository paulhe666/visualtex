//! P08 Matrix foundation acceptance. Real core actions plus fresh-parser oracles;
//! these tests do not claim rectangular selection, Cases, Aligned or Office UI.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{json,Value};
use visualtex_engine_core::{
    editing::{EditorAction,EditorActionResult},fonts::FontRegistry,grids::GridKind,
    layout::{LayoutEngine,LayoutOptions},model::{self,IdAllocator,LatexMode,Node,NodeKind},
    parser,positions::{Horizontal,Position,Selection,TreeIndex},session::{EngineSession,NodeId},
    source::SourceText,transactions::{EditCommand,InputMode},
};
struct Owned { session:EngineSession, sequence:u32 }
impl Owned {
    fn new(source:&str)->Self {Self{session:EngineSession::from_source("matrix-review".into(),source.into()).unwrap(),sequence:0}}
    fn state(&self)->Value {serde_json::to_value(self.session.editor_snapshot().unwrap()).unwrap()}
    fn action(&mut self,action:EditorAction)->EditorActionResult {
        let before=self.session.document_snapshot();self.sequence+=1;
        let result=self.session.apply_action(self.session.revision(),self.sequence,action).unwrap();
        let index=TreeIndex::new(&result.editor.document.analysis.tree).unwrap();
        index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map)=&result.position_map {
            for position in TreeIndex::new(&before.analysis.tree).unwrap().carets {
                index.validate_position(map.apply_position(position,&index).unwrap()).unwrap();
            }
        }
        let source=&result.editor.document.source;
        let parsed=parser::parse(SourceText::new(1,source.as_str()).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap();
        assert!(parsed.complete&&parsed.supported,"{}: {:?}",source,parsed.diagnostics);
        assert_eq!(semantic(serde_json::to_value(&parsed.tree).unwrap()),semantic(serde_json::to_value(&result.editor.document.analysis.tree).unwrap()));
        result
    }
    fn edit(&mut self,command:EditCommand)->EditorActionResult {self.action(EditorAction::Edit{command,group_id:None})}
    fn select(&mut self,position:Position){self.action(EditorAction::SetSelection{selection:Selection::caret(position)});}
    fn cell(&mut self,row:usize,column:usize,offset:usize){
        let snapshot=self.session.document_snapshot();let id=cell(first_matrix(&snapshot.analysis.tree),row,column).id;
        self.select(Position::row(id,offset));
    }
    fn reject(&mut self,command:EditCommand){
        let before=self.state();self.sequence+=1;
        assert!(self.session.apply_action(self.session.revision(),self.sequence,EditorAction::Edit{command,group_id:None}).is_err());
        assert_eq!(self.state(),before);
    }
}
fn semantic(value:Value)->Value {match value {
    Value::Object(map)=>Value::Object(map.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"columns"|"leading"|"trailing"|"beforeBeginName"|"beforeEndName"))
        .map(|(k,v)|(k,semantic(v))).collect()),
    Value::Array(a)=>Value::Array(a.into_iter().map(semantic).collect()),x=>x,
}}
fn content(v:&Value)->Value {json!({"source":v["document"]["source"],"tree":v["document"]["analysis"]["tree"],"selection":v["selection"],"documentId":v["document"]["documentId"]})}
fn first_matrix(node:&Node)->&Node {
    if matches!(node.kind,NodeKind::Grid{environment:GridKind::Matrix,..}){return node;}
    for child in node.children(){let mut found=false;child.visit(&mut|n|found|=matches!(n.kind,NodeKind::Grid{environment:GridKind::Matrix,..}));if found{return first_matrix(child);}}
    panic!("expected matrix")
}
fn cell(grid:&Node,row:usize,column:usize)->&Node {
    let NodeKind::Grid{rows,..}=&grid.kind else{panic!("not grid")};
    let NodeKind::GridRow{cells,..}=&rows[row].kind else{panic!("not row")};&cells[column]
}
fn shape(grid:&Node)->(usize,usize){let NodeKind::Grid{rows,columns,..}=&grid.kind else{panic!("not grid")};(rows.len(),columns.len())}
fn ids(node:&Node)->BTreeSet<u32>{let mut out=BTreeSet::new();node.visit(&mut|n|{assert!(out.insert(n.id.0));if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(out.insert(c.0));}}});out}
fn matrix_source(rows:usize,columns:usize)->String {
    format!("\\begin{{matrix}}{}\\end{{matrix}}",(0..rows).map(|r|(0..columns).map(|c|if (r+c)%5==0{String::new()}else{format!("a_{{{r}{c}}}")}).collect::<Vec<_>>().join("&")).collect::<Vec<_>>().join("\\\\"))
}
#[test]
fn rectangular_sizes_empty_cells_identity_and_source_roundtrip(){
    for (r,c) in [(1,1),(2,2),(3,3),(16,16),(32,12)] {
        let source=matrix_source(r,c);let t=Owned::new(&source);let s=t.session.document_snapshot();
        assert!(s.analysis.complete&&s.analysis.supported);assert_eq!(shape(first_matrix(&s.analysis.tree)),(r,c));
        assert_eq!(model::latex(&s.analysis.tree,LatexMode::Lossless),source);
        assert_eq!(ids(&s.analysis.tree).len(),s.analysis.facts.nodes+s.analysis.facts.grid_columns);
        let index=TreeIndex::new(&s.analysis.tree).unwrap();
        for row in 0..r {for column in 0..c {index.validate_position(Position::row(cell(first_matrix(&s.analysis.tree),row,column).id,0)).unwrap();}}
    }
}
#[test]
fn separators_are_scoped_to_tokens_and_nearest_environment(){
    for source in [
        "\\begin %begin\n{matrix}\\text{A\\&B}&\\frac{x}{y}\\\\%hidden & \\\\\n\\sqrt{z}&\\begin{matrix}a&b\\\\c&d\\end{matrix}\\end %end\n{matrix}",
        r"\begin{matrix}\left[a\right]&\&\\\color{red}x&\text{中文}\end{matrix}",
        r"\begin{matrix}[a]&b\\c&d\end{matrix}",
    ] {
        let mut t=Owned::new(source);let s=t.session.document_snapshot();assert!(s.analysis.supported,"{:?}",s.analysis.diagnostics);
        assert_eq!(shape(first_matrix(&s.analysis.tree)),(2,2));assert_eq!(t.session.source(),source);
        t.cell(0,1,0);t.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math});
        assert!(t.session.source().contains("\\begin"));
    }
}
#[test]
fn ragged_options_and_unknown_environments_remain_preserved_and_unsupported(){
    for source in [r"\begin{matrix}a&b\\c\end{matrix}",r"\begin{matrix}a&b\\[2pt]c&d\end{matrix}",r"\begin{cases}a&b&c\end{cases}",r"\begin{aligned}[t]a&b\end{aligned}"]{
        let t=Owned::new(source);let s=t.session.document_snapshot();assert_eq!(t.session.source(),source);assert!(!s.analysis.supported&&!s.analysis.exportable);
    }
}
#[test]
fn insert_matrix_moves_selected_math_without_reidentifying_and_is_one_undo(){
    let mut t=Owned::new("xy");let before=t.session.document_snapshot();
    let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};let old_ids:Vec<_>=entries.iter().map(|e|e.node.id).collect();
    t.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(before.analysis.tree.id,0),focus:Position::row(before.analysis.tree.id,2)}});
    let initial=t.state();let out=t.edit(EditCommand::InsertMatrix{rows:2,columns:2});let after=t.state();
    assert_eq!(out.editor.undo_depth,1);let grid=first_matrix(&out.editor.document.analysis.tree);assert_eq!(shape(grid),(2,2));
    let NodeKind::Row{entries,..}=&cell(grid,0,0).kind else{panic!()};assert_eq!(entries.iter().map(|e|e.node.id).collect::<Vec<_>>(),old_ids);
    t.action(EditorAction::Undo);assert_eq!(content(&t.state()),content(&initial));
    t.action(EditorAction::Redo);assert_eq!(content(&t.state()),content(&after));
}
#[test]
fn row_insertion_deletion_preserves_survivor_and_column_ids_with_exact_history(){
    let mut t=Owned::new(r"\begin{matrix}a&b\\c&d\end{matrix}");t.cell(0,1,1);let before=t.state();
    let old=t.session.document_snapshot();let old_grid=first_matrix(&old.analysis.tree);let old_ids=ids(old_grid);
    let out=t.edit(EditCommand::InsertGridRow{after:true});let grid=first_matrix(&out.editor.document.analysis.tree);
    assert_eq!(shape(grid),(3,2));assert!(old_ids.is_subset(&ids(grid)));assert_eq!(cell(grid,2,0).id,cell(old_grid,1,0).id);
    assert_eq!(out.editor.selection,Selection::caret(Position::row(cell(grid,1,1).id,0)));let inserted=t.state();
    t.action(EditorAction::Undo);assert_eq!(content(&t.state()),content(&before));t.action(EditorAction::Redo);assert_eq!(content(&t.state()),content(&inserted));
    let out=t.edit(EditCommand::DeleteGridRow);assert_eq!(out.editor.document.source,old.source);assert_eq!(ids(first_matrix(&out.editor.document.analysis.tree)),old_ids);
}
#[test]
fn columns_insert_before_and_after_and_delete_preserve_independent_cell_ids(){
    for after in [false,true] {
        let mut t=Owned::new(r"\begin{matrix}\alpha &b\\c&\frac{x}{y}\end{matrix}");t.cell(0,0,0);let old=t.session.document_snapshot();let original=ids(first_matrix(&old.analysis.tree));
        let out=t.edit(EditCommand::InsertGridColumn{after});let grid=first_matrix(&out.editor.document.analysis.tree);assert_eq!(shape(grid),(2,3));assert!(original.is_subset(&ids(grid)));
        let inserted=t.state();let out=t.edit(EditCommand::DeleteGridColumn);assert_eq!(out.editor.document.source,old.source);assert_eq!(ids(first_matrix(&out.editor.document.analysis.tree)),original);
        t.action(EditorAction::Undo);assert_eq!(content(&t.state()),content(&inserted));
    }
}
#[test]
fn shift_enter_inside_fraction_adds_row_without_splitting_math(){
    let mut t=Owned::new(r"\begin{matrix}a&\frac{xy}{z}\\c&d\end{matrix}");let old=t.session.document_snapshot();let grid=first_matrix(&old.analysis.tree);let saved=serde_json::to_value(cell(grid,0,1)).unwrap();
    let mut numerator=None;cell(grid,0,1).visit(&mut|n|if let NodeKind::Fraction{numerator:n,..}=&n.kind{numerator=Some(n.id)});
    t.select(Position::row(numerator.unwrap(),1));let out=t.edit(EditCommand::SplitLine);let next=first_matrix(&out.editor.document.analysis.tree);
    assert_eq!(shape(next),(3,2));assert_eq!(out.editor.document.analysis.facts.grids,1);assert_eq!(ids(cell(next,0,1)),ids(cell(grid,0,1)));
    assert_eq!(semantic(serde_json::to_value(cell(next,0,1)).unwrap()),semantic(saved));
    assert_eq!(out.editor.selection,Selection::caret(Position::row(cell(next,1,1).id,0)));
}
#[test]
fn nested_matrix_and_gathered_route_to_the_nearest_environment(){
    let mut t=Owned::new(r"\begin{matrix}a&\begin{matrix}x&y\\z&w\end{matrix}\\b&c\end{matrix}");
    let old=t.session.document_snapshot();let outer=first_matrix(&old.analysis.tree);let inner=first_matrix(cell(outer,0,1));let inner_id=inner.id;
    t.select(Position::row(cell(inner,0,1).id,1));let out=t.edit(EditCommand::SplitLine);let outer=first_matrix(&out.editor.document.analysis.tree);
    assert_eq!(shape(outer),(2,2));assert_eq!(shape(out.editor.document.analysis.tree.find(inner_id).unwrap()),(3,2));assert_eq!(out.editor.document.analysis.facts.grids,2);
    let mut t=Owned::new(r"\begin{matrix}\begin{gathered}x\\y\end{gathered}&z\end{matrix}");let old=t.session.document_snapshot();let outer=first_matrix(&old.analysis.tree);
    let mut position=None;cell(outer,0,0).visit(&mut|n|if let NodeKind::Grid{environment:GridKind::Gathered,rows,..}=&n.kind{let NodeKind::GridRow{cells,..}=&rows[0].kind else{panic!()};position=Some(Position::row(cells[0].id,1));});
    t.select(position.unwrap());let out=t.edit(EditCommand::SplitLine);assert_eq!(shape(first_matrix(&out.editor.document.analysis.tree)),(1,2));assert_eq!(out.editor.document.analysis.facts.grids,2);assert_eq!(out.editor.document.analysis.facts.grid_rows,4);
}
#[test]
fn cell_navigation_and_delete_boundaries_never_merge_cells_or_create_history(){
    let mut t=Owned::new(r"\begin{matrix}a&\\c&d\end{matrix}");t.cell(0,0,0);let initial=t.session.document_snapshot();
    for direction in [Horizontal::Right,Horizontal::Right,Horizontal::Left] {let out=t.edit(EditCommand::MoveGridCell{direction});assert_eq!(out.editor.document.revision,initial.revision);assert_eq!(out.editor.undo_depth,0);}
    let state=t.session.document_snapshot();let grid=first_matrix(&state.analysis.tree);assert_eq!(t.session.editor_snapshot().unwrap().selection,Selection::caret(Position::row(cell(grid,0,1).id,0)));
    for (r,c,offset,command) in [(0,0,0,EditCommand::DeleteBackward),(0,0,1,EditCommand::DeleteForward),(0,1,0,EditCommand::DeleteBackward),(1,1,1,EditCommand::DeleteForward)]{
        t.cell(r,c,offset);let before=t.session.document_snapshot();let out=t.edit(command);assert_eq!(out.editor.document.source,before.source);assert_eq!(out.editor.document.revision,before.revision);assert_eq!(out.editor.undo_depth,0);
    }
}
#[test]
fn invalid_dimension_and_last_row_or_column_deletion_are_atomic(){
    let mut t=Owned::new("x");for (rows,columns) in [(0,1),(1,0),(257,1),(1,65),(256,64),(u32::MAX,u32::MAX)]{t.reject(EditCommand::InsertMatrix{rows,columns});}
    let mut t=Owned::new(r"\begin{matrix}x\end{matrix}");t.cell(0,0,0);t.reject(EditCommand::DeleteGridRow);t.reject(EditCommand::DeleteGridColumn);
}
#[test]
fn whole_matrix_copy_has_disjoint_node_and_column_ids(){
    let mut t=Owned::new(r"\begin{matrix}a&b\\c&d\end{matrix}");let old=t.session.document_snapshot();let grid=first_matrix(&old.analysis.tree);let original=ids(grid);
    t.action(EditorAction::SetSelection{selection:Selection::Node{node_id:grid.id}});let out=t.edit(EditCommand::DuplicateSelection);
    let NodeKind::Row{entries,..}=&out.editor.document.analysis.tree.kind else{panic!()};assert_eq!(entries.len(),2);assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));assert_eq!(ids(&entries[0].node),original);
}
#[test]
fn partial_matrix_selection_cannot_silently_delete_or_copy_the_entire_grid(){
    let mut t=Owned::new(r"\begin{matrix}a&b\\c&d\end{matrix}");
    let old=t.session.document_snapshot();let grid=first_matrix(&old.analysis.tree);
    t.action(EditorAction::SetSelection{selection:Selection::Range{
        anchor:Position::row(cell(grid,0,0).id,0),focus:Position::row(cell(grid,1,1).id,1)}});
    let before=t.state();
    assert!(t.session.copy_selection(t.session.revision()).is_err());assert_eq!(t.state(),before);
    for command in [EditCommand::DeleteBackward,EditCommand::DeleteForward,EditCommand::DuplicateSelection,
        EditCommand::InsertText{text:"x".into(),mode:InputMode::Math},EditCommand::PasteLatex{latex:"y".into()}] {
        t.reject(command);
    }
    t.cell(0,0,0);t.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math});
    assert!(t.session.source().contains("qa&b"));
}
#[test]
fn matrix_shared_columns_cell_hit_targets_and_natural_svg_are_consistent(){
    let t=Owned::new(r"\begin{matrix}a&\frac{long}{x}&\\\sqrt{b}&y&c\\d&e&f\end{matrix}");
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;
    let mut options=LayoutOptions::new(id);let mut layout=LayoutEngine::default();let scene=layout.layout(&t.session,&fonts,&options).unwrap();scene.validate().unwrap();
    assert_eq!(scene.grid_columns.len(),3);assert!(scene.grid_columns.windows(2).all(|w|w[0].x+w[0].width<w[1].x));
    let state=t.session.document_snapshot();let grid=first_matrix(&state.analysis.tree);
    for row in 0..3 {for column in 0..3 {let node=cell(grid,row,column);let caret=scene.caret(Position::row(node.id,0)).unwrap();let anchor=&scene.grid_columns[column];assert!(caret.x>=anchor.x-1e-9&&caret.x<=anchor.x+anchor.width+1e-9);assert_eq!(caret.baseline,scene.caret(Position::row(cell(grid,row,0).id,0)).unwrap().baseline);}}
    let empty=Position::row(cell(grid,0,2).id,0);let caret=scene.caret(empty).unwrap();assert!(scene.hit_test(caret.x,caret.baseline).unwrap().same_location(empty));
    let svg=scene.svg().unwrap();assert_eq!(svg.matches("<svg").count(),1);assert!(!svg.contains("<text")&&!svg.contains("foreignObject"));
    options.font_size_px*=2.0;let doubled=layout.layout(&t.session,&fonts,&options).unwrap();assert_eq!(doubled.width,scene.width*2.0);assert_eq!(doubled.height,scene.height*2.0);assert_eq!(doubled.rebuilt_fragments,0);assert_eq!(serde_json::to_value(doubled.grid_columns).unwrap(),serde_json::to_value(scene.grid_columns).unwrap());
}
