//! P07.6 regression: explicit glue, invisible mathematical boxes, editable
//! enclosures and source/structure agreement. No renderer fallback or Office IO.
use std::{collections::BTreeSet,path::PathBuf};
use visualtex_engine_core::{ErrorCode,editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,IdAllocator,LatexMode,Node,NodeKind},
    parser::{self,ParseResult},positions::{Position,Selection,TreeIndex,Horizontal},scene::LayoutSnapshot,
    session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
    transactions::{EditCommand,InputMode},spacing};

const SPACES:&[&str]=&[r"\,",r"\:",r"\;",r"\!",r"\thinspace",r"\medspace",r"\thickspace",
    r"\negthinspace",r"\negmedspace",r"\negthickspace",r"\enspace",r"\enskip",r"\quad",r"\qquad"];
fn parse(s:&str)->ParseResult {parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn ids(n:&Node)->BTreeSet<NodeId>{let mut ids=BTreeSet::new();n.visit(&mut |n|assert!(ids.insert(n.id)));ids}
fn fonts()->(FontRegistry,LayoutOptions){
    let p=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(p).unwrap(),0).unwrap().id;
    (fonts,LayoutOptions::new(id))
}
fn draw(s:&str,style:MathStyle)->LayoutSnapshot {
    let (fonts,mut opts)=fonts();opts.style=style;
    let session=EngineSession::from_source("space-geometry".into(),s.into()).unwrap();
    LayoutEngine::default().layout(&session,&fonts,&opts).unwrap()
}
fn close(a:f64,b:f64){assert!((a-b).abs()<1e-9,"{a} != {b}");}
fn assert_projection(doc:&visualtex_engine_core::session::DocumentSnapshot){
    let fresh=parse(&doc.source);
    assert_eq!(model::latex(&doc.analysis.tree,LatexMode::Canonical),model::latex(&fresh.tree,LatexMode::Canonical),"{}",doc.source);
    assert_eq!(doc.analysis.facts,fresh.facts,"{}",doc.source);
    assert_eq!(ids(&doc.analysis.tree).len(),doc.analysis.facts.nodes);
    doc.analysis.tree.visit(&mut |n| {
        let o=n.origin.as_ref().expect("Every projected node has an origin");
        assert_eq!(o.span.source_id,doc.source_id);
        assert_eq!(o.text(),&doc.source[o.span.byte_start as usize..o.span.byte_end as usize]);
    });
}
struct Owned{session:EngineSession,sequence:u32}
impl Owned{
    fn new(source:&str)->Self {Self{session:EngineSession::from_source("spacing-edit".into(),source.into()).unwrap(),sequence:0}}
    fn action(&mut self,a:EditorAction)->EditorActionResult {
        let before=self.session.document_snapshot();self.sequence+=1;
        let r=self.session.apply_action(self.session.revision(),self.sequence,a).unwrap();
        let index=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();index.validate_selection(&r.editor.selection).unwrap();
        if let Some(map)=&r.position_map {for p in TreeIndex::new(&before.analysis.tree).unwrap().carets {map.apply_position(p,&index).unwrap();}}
        assert_projection(&r.editor.document);r
    }
    fn edit(&mut self,c:EditCommand)->EditorActionResult {self.action(EditorAction::Edit{command:c,group_id:None})}
    fn all(&mut self){let d=self.session.document_snapshot();let NodeKind::Row{entries,..}=&d.analysis.tree.kind else{panic!()};
        self.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(d.analysis.tree.id,0),focus:Position::row(d.analysis.tree.id,entries.len())}});}
}
#[test]
fn explicit_spaces_preserve_spelling_and_are_not_ordinary_math_symbols(){
    for cmd in SPACES {
        let source=format!("x{cmd} % retained\r\n +y");let p=parse(&source);
        assert!(p.complete&&p.supported&&p.exportable,"{source}");assert_eq!(p.facts.spaces,1);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
        let q=parse(&model::latex(&p.tree,LatexMode::Canonical));assert!(q.exportable);assert_eq!(q.facts,p.facts);
    }
}
#[test]
fn phantom_and_boxed_sources_preserve_empty_unknown_and_unclosed_states(){
    for cmd in [r"\phantom",r"\hphantom",r"\vphantom",r"\boxed"] {
        for suffix in ["", "{"]{let s=format!("{cmd}{suffix}");let p=parse(&s);assert!(!p.complete&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),s);}
        let s=format!("{cmd}{{}}");let p=parse(&s);assert!(p.complete&&p.supported&&!p.exportable,"{s}");
        let s=format!("{cmd}{{\\future{{x}}}}");let p=parse(&s);assert!(!p.supported&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),s);
        let s=format!("{cmd} % author\n {{\\frac{{x}}{{y}}}}");let p=parse(&s);assert!(p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),s);
    }
}
#[test]
fn phantom_has_exact_logical_dimensions_without_any_hidden_paint(){
    let visible=draw(r"\frac{x}{y}",MathStyle::Display);
    let original=&visible.regions[0].logical;
    for (cmd,keep_width,keep_height) in [(r"\phantom",true,true),(r"\hphantom",true,false),(r"\vphantom",false,true)] {
        let s=draw(&format!("{cmd}{{\\frac{{x}}{{y}}}}"),MathStyle::Display);
        let region=s.regions.iter().find(|r|r.role=="phantom").unwrap();
        close(region.logical.width(),if keep_width {original.width()}else{0.0});
        close(region.logical.height(),if keep_height {original.height()}else{0.0});
        assert!(s.ink.is_none()&&s.glyphs.is_empty()&&s.rules.is_empty()&&s.resources.is_empty());
        assert!(s.regions.iter().all(|r|r.ink.is_none()));
        assert!(s.carets.len()>4,"Hidden arguments must retain real editing positions");
        let svg=s.svg().unwrap();assert!(!svg.contains("<path")&&!svg.contains("<rect")&&!svg.contains("opacity"));
    }
}
#[test]
fn nested_hidden_geometry_does_not_remove_visible_sibling_or_outer_decoration(){
    let source=r"x+\phantom{\boxed{\frac{a}{b}}}+\boxed{y}";
    let doc=EngineSession::from_source("hidden-subtree".into(),source.into()).unwrap().document_snapshot();
    let NodeKind::Row{entries,..}=&doc.analysis.tree.kind else{panic!()};
    let hidden=ids(&entries[2].node);let s=draw(source,MathStyle::Display);
    assert!(s.glyphs.iter().all(|g|!hidden.contains(&g.node_id)));
    assert!(s.rules.iter().all(|r|!hidden.contains(&r.node_id)));
    assert_eq!(s.rules.len(),4);assert!(!s.glyphs.is_empty());
    for id in hidden {assert!(s.regions.iter().any(|r|r.node_id==id));}
}
#[test]
fn explicit_glue_preserves_binary_classification_and_adds_only_authored_distance(){
    let baseline=draw("x+y",MathStyle::Display).advance_em;
    close(draw(r"x\,+y",MathStyle::Display).advance_em-baseline,3.0/18.0);
    close(draw(r"x+\,y",MathStyle::Display).advance_em-baseline,3.0/18.0);
    close(draw(r"x\!+y",MathStyle::Display).advance_em-baseline,-3.0/18.0);
    close(draw(r"x\!\!+\,y",MathStyle::Display).advance_em-baseline,-3.0/18.0);
    close(draw(r"\,+x",MathStyle::Display).advance_em-draw("+x",MathStyle::Display).advance_em,3.0/18.0);
    close(draw(r"x+\,",MathStyle::Display).advance_em-draw("x+",MathStyle::Display).advance_em,3.0/18.0);
}
#[test]
fn negative_glue_can_move_carets_back_without_invalid_or_clipped_export_bounds(){
    let s=draw(r"\!\!",MathStyle::Display);close(s.advance_em,-6.0/18.0);s.validate().unwrap();
    assert!(s.width>0.0&&s.carets.windows(2).any(|p|p[1].x<p[0].x));
    for source in [r"\!x",r"x\!\!\!\!\!",r"\frac{\!}{\!}",r"\boxed{\!}",r"\phantom{\!}x"] {
        let s=draw(source,MathStyle::Display);s.validate().unwrap();s.svg().unwrap();
    }
}
#[test]
fn mu_spaces_follow_script_style_while_em_skips_keep_surrounding_size(){
    let (fonts,_)=fonts();let id=fonts.len();assert_eq!(id,1);
    let thin=draw(r"\,",MathStyle::Display).advance_em;
    let script=draw(r"\,",MathStyle::Script).advance_em;assert!(script<thin&&script>0.0);
    close(draw(r"\quad",MathStyle::Display).advance_em,draw(r"\quad",MathStyle::ScriptScript).advance_em);
    close(draw(r"\qquad",MathStyle::Script).advance_em,2.0);
}
#[test]
fn boxed_uses_display_body_four_vector_rules_and_not_editor_css(){
    let s=draw(r"\boxed{\frac{x}{y}}",MathStyle::Script);
    let normal=draw(r"\frac{x}{y}",MathStyle::Display);
    assert_eq!(s.rules.len(),5);assert_eq!(s.glyphs.len(),normal.glyphs.len());
    assert_eq!(s.glyphs.iter().map(|g|g.scale).collect::<Vec<_>>(),normal.glyphs.iter().map(|g|g.scale).collect::<Vec<_>>());
    let svg=s.svg().unwrap();assert_eq!(svg.matches("<rect ").count(),5);assert!(!svg.contains("stroke="));
    assert!(s.width>normal.width&&s.height>normal.height);
}
#[test]
fn wrapping_copy_edit_delete_and_history_restore_real_hidden_node_identities(){
    for command in [EditCommand::InsertPhantom{command:r"\phantom".into()},EditCommand::InsertBoxed] {
        let mut t=Owned::new("x+y");let original=ids(&t.session.document_snapshot().analysis.tree);t.all();t.edit(command);
        let wrapped=t.session.document_snapshot();assert!(original.is_subset(&ids(&wrapped.analysis.tree)));
        t.all();t.edit(EditCommand::DuplicateSelection);let duplicated=t.session.document_snapshot();
        let NodeKind::Row{entries,..}=&duplicated.analysis.tree.kind else{panic!()};assert_eq!(entries.len(),2);
        assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));
        t.edit(EditCommand::DeleteBackward);assert_eq!(t.session.source(),wrapped.source);
        t.action(EditorAction::Undo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&duplicated.analysis.tree));
        t.action(EditorAction::Redo);assert_eq!(t.session.source(),wrapped.source);
    }
}
#[test]
fn blank_wrappers_accept_typing_and_exit_without_a_content_change(){
    for command in [EditCommand::InsertPhantom{command:r"\vphantom".into()},EditCommand::InsertBoxed] {
        let mut t=Owned::new("");let a=t.edit(command);assert!(!a.editor.document.analysis.exportable);
        t.edit(EditCommand::InsertText{text:"x".into(),mode:InputMode::Math});let before=t.session.revision();
        t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});assert_eq!(t.session.revision(),before);
        t.edit(EditCommand::InsertText{text:"+y".into(),mode:InputMode::Math});assert!(t.session.source().ends_with("}+y"));
    }
}
#[test]
fn space_insert_delete_is_one_edit_and_unsupported_insertion_is_atomic(){
    let mut t=Owned::new("");t.edit(EditCommand::InsertSpace{command:r"\!".into()});assert_eq!(t.session.source(),r"\!");
    t.edit(EditCommand::DeleteBackward);assert_eq!(t.session.source(),"");t.action(EditorAction::Undo);assert_eq!(t.session.source(),r"\!");
    let before=serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap();
    for command in [EditCommand::InsertSpace{command:r"\infiniteglue".into()},EditCommand::InsertPhantom{command:r"\hidefuture".into()}] {
        let e=t.session.apply_action(t.session.revision(),t.sequence+1,EditorAction::Edit{command,group_id:None}).unwrap_err();
        assert_eq!(e.code,ErrorCode::UnsupportedSyntax);assert_eq!(serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap(),before);
    }
}
#[test]
fn patching_nested_new_wrappers_updates_actual_math_and_rebases_every_origin(){
    for cmd in [r"\mathbf",r"\phantom",r"\hphantom",r"\vphantom",r"\boxed"] {
        let mut s=EngineSession::from_source("new-wrapper-patch".into(),format!("{cmd}{{\\frac{{x}}{{z}}}}+{cmd}{{x}}" )).unwrap();
        let before=s.document_snapshot();let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};let sibling=ids(&entries[2].node);
        let at=s.source().find('x').unwrap();s.patch_source(s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:(at+1)as u32,replacement:"yy".into()}).unwrap();
        let after=s.document_snapshot();assert_projection(&after);let NodeKind::Row{entries,..}=&after.analysis.tree.kind else{panic!()};
        assert_eq!(ids(&entries[2].node),sibling);
    }
}
#[test]
fn seeded_new_wrapper_source_edits_match_a_fresh_parse_without_losing_drafts(){
    let mut seed=0x706e_2026u32;
    let initial=r"\boxed{\phantom{\mathbf{x}}+\frac{a\!b}{\hphantom{y}}}+\quad z";
    let replacements=["q", "", "{", "}", r"\!", "中", "%\n", r"\quad "];
    let mut s=EngineSession::from_source("seeded-spacing".into(),initial.into()).unwrap();
    for i in 0..480 {
        if i%16==0 {s.replace_source(s.revision(),initial.into()).unwrap();}
        seed=seed.wrapping_mul(1664525).wrapping_add(1013904223);
        let boundaries:Vec<_>=s.source().char_indices().map(|(b,_)|b).chain(std::iter::once(s.source().len())).collect();
        let pick=seed as usize%boundaries.len();let start=boundaries[pick];let end=boundaries[(pick+usize::from(seed&1==0)).min(boundaries.len()-1)];
        let start_utf16=s.source()[..start].encode_utf16().count()as u32;let end_utf16=s.source()[..end].encode_utf16().count()as u32;
        let patch=SourceEdit{start_utf16,end_utf16,replacement:replacements[(seed>>8)as usize%replacements.len()].into()};
        let prior=s.document_snapshot();match s.patch_source(s.revision(),patch) {
            Ok(_)=>{let d=s.document_snapshot();assert_projection(&d);let p=parse(&d.source);
                assert_eq!((d.analysis.complete,d.analysis.supported,d.analysis.exportable),(p.complete,p.supported,p.exportable),"seed={seed} i={i}");},
            Err(_)=>assert_eq!(serde_json::to_value(s.document_snapshot()).unwrap(),serde_json::to_value(prior).unwrap()),
        }
    }
}
#[test]
fn transparent_layout_changes_and_hidden_cache_reuse_do_not_poison_visible_reuse(){
    let (f,mut options)=fonts();let mut layout=LayoutEngine::default();
    let mut s=EngineSession::from_source("phantom-cache".into(),r"\phantom{\frac{x}{y}}".into()).unwrap();
    let a=layout.layout(&s,&f,&options).unwrap();assert!(a.glyphs.is_empty());
    options.font_size_px*=2.0;let b=layout.layout(&s,&f,&options).unwrap();close(b.width,a.width*2.0);assert_eq!(b.rebuilt_fragments,0);
    s.replace_source(s.revision(),r"\frac{x}{y}".into()).unwrap();let c=layout.layout(&s,&f,&options).unwrap();assert!(!c.glyphs.is_empty()&&!c.rules.is_empty());
    assert_eq!(spacing::phantom_axis(r"\vphantom"),Some(spacing::PhantomAxis::Vertical));
}
