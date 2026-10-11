use std::path::PathBuf;
use serde_json::{Value,json};
use visualtex_engine_core::Runtime;
fn send(runtime:&mut Runtime,command:Value)->Value {
    serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
}
fn ok(runtime:&mut Runtime,command:Value)->Value {
    let result=send(runtime,command);assert_eq!(result["status"],"ok","{result}");result["result"].clone()
}
fn setup(source:&str)->(Runtime,Value) {
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut runtime=Runtime::default();let font=runtime.register_font(&std::fs::read(path).unwrap(),0).unwrap();
    ok(&mut runtime,json!({"op":"create","documentId":"queries","source":source}));
    (runtime,json!({"mathFont":font.id,"textFont":null,"fontSizePx":16,"paddingEm":0.125,"style":"display"}))
}
fn layout(runtime:&mut Runtime,options:&Value,revision:u32)->Value {
    ok(runtime,json!({"op":"layout","handle":1,"expectedRevision":revision,"options":options,"includeSvg":false}))
}
fn query(runtime:&mut Runtime,token:&Value,revision:u32,query:Value)->Value {
    ok(runtime,json!({"op":"queryScene","handle":1,"expectedRevision":revision,"sceneToken":token,"query":query}))["result"].clone()
}
fn row(id:&Value,offset:u32)->Value {json!({"kind":"row","nodeId":id,"offset":offset,"affinity":"after"})}
fn text(id:&Value,offset:u32)->Value {json!({"kind":"text","nodeId":id,"grapheme":offset,"affinity":"after"})}
#[test]
fn geometry_queries_never_edit_the_document_or_advance_its_revision() {
    let (mut runtime,options)=setup(r"\frac{a+b}{c+d}");
    let before=ok(&mut runtime,json!({"op":"inspectEditor","handle":1}));
    let scene=layout(&mut runtime,&options,0);let token=&scene["sceneToken"];
    let fraction=&before["snapshot"]["document"]["analysis"]["tree"]["entries"][0]["node"];
    let from=row(&fraction["numerator"]["id"],2);
    let result=query(&mut runtime,token,0,json!({"kind":"vertical","from":from,"down":true,"desiredX":null}));
    assert_eq!(result["position"]["nodeId"],fraction["denominator"]["id"]);
    for _ in 0..100 {let r=query(&mut runtime,token,0,json!({"kind":"hitTest","xEm":0.4,"yEm":0.5}));assert!(r["position"]["nodeId"].is_number());}
    let after=ok(&mut runtime,json!({"op":"inspectEditor","handle":1}));assert_eq!(before,after);
}
#[test]
fn font_size_changes_retire_old_geometry_tokens_even_at_the_same_revision() {
    let (mut runtime,mut options)=setup("x+1");let a=layout(&mut runtime,&options,0);
    options["fontSizePx"]=json!(32);let b=layout(&mut runtime,&options,0);
    assert_ne!(a["sceneToken"],b["sceneToken"]);
    let result=send(&mut runtime,json!({"op":"queryScene","handle":1,"expectedRevision":0,"sceneToken":a["sceneToken"],"query":{"kind":"hitTest","xEm":0,"yEm":0}}));
    assert_eq!(result["error"]["code"],"conflict");
    query(&mut runtime,&b["sceneToken"],0,json!({"kind":"hitTest","xEm":0,"yEm":0}));
}
#[test]
fn editing_requires_new_geometry_but_selection_only_changes_do_not() {
    let (mut runtime,options)=setup("x+1");let scene=layout(&mut runtime,&options,0);let token=&scene["sceneToken"];
    let p=row(&json!(1),1);
    ok(&mut runtime,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":1,"action":{"kind":"setSelection","selection":{"kind":"range","anchor":p,"focus":p}}}));
    let geometry=query(&mut runtime,token,0,json!({"kind":"selection","selection":{"kind":"range","anchor":p,"focus":p}}));
    assert_eq!(geometry["caret"],p);assert_eq!(geometry["rectangles"].as_array().unwrap().len(),0);
    ok(&mut runtime,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":2,"action":{"kind":"edit","command":{"kind":"insertText","text":"y","mode":"math"},"groupId":null}}));
    for (revision,error) in [(0,"staleRevision"),(1,"conflict")] {
        let result=send(&mut runtime,json!({"op":"queryScene","handle":1,"expectedRevision":revision,"sceneToken":token,"query":{"kind":"hitTest","xEm":0,"yEm":0}}));
        assert_eq!(result["error"]["code"],error);
    }
    let new_scene=layout(&mut runtime,&options,1);query(&mut runtime,&new_scene["sceneToken"],1,json!({"kind":"hitTest","xEm":0,"yEm":0}));
}
#[test]
fn text_selection_uses_grapheme_edges_and_reports_structural_location() {
    let (mut runtime,options)=setup("\\frac{\\text{Ae\u{301}B}}{x}");let scene=layout(&mut runtime,&options,0);
    let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
    let text_id=&doc["snapshot"]["analysis"]["tree"]["entries"][0]["node"]["numerator"]["entries"][0]["node"]["id"];
    let result=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"selection","selection":{"kind":"range","anchor":text(text_id,1),"focus":text(text_id,2)}}));
    assert!(result["caret"].is_null());assert_eq!(result["rectangles"].as_array().unwrap().len(),1);
    assert_eq!(result["description"]["path"],json!(["root","fraction","numerator","text"]));
    assert_eq!(result["description"]["length"],3);assert_eq!(result["description"]["offset"],2);
    let r=&result["rectangles"][0];assert!(r["xMax"].as_f64().unwrap()>r["xMin"].as_f64().unwrap());
}
#[test]
fn whole_structure_selection_is_a_single_rectangle_with_no_caret() {
    let (mut runtime,options)=setup(r"\frac{x}{y}+1");let scene=layout(&mut runtime,&options,0);
    let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
    let id=&doc["snapshot"]["analysis"]["tree"]["entries"][0]["node"]["id"];
    let result=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"selection","selection":{"kind":"node","nodeId":id}}));
    assert!(result["caret"].is_null());assert_eq!(result["rectangles"].as_array().unwrap().len(),1);
}
#[test]
fn decoration_geometry_describes_editable_label_and_nested_accent_carets() {
    let (mut runtime,options)=setup(r"\overset{a}{\hat{\vec{x}}}");
    let scene=layout(&mut runtime,&options,0);
    let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
    let annotation=&doc["snapshot"]["analysis"]["tree"]["entries"][0]["node"];
    let hat=&annotation["base"]["entries"][0]["node"];
    let vec=&hat["body"]["entries"][0]["node"];
    let cases=[
        (row(&annotation["annotation"]["id"],1),json!(["root","annotation","annotationLabel"])),
        (row(&annotation["base"]["id"],0),json!(["root","annotation","annotationBase"])),
        (row(&vec["body"]["id"],1),json!(["root","annotation","annotationBase","accent","accentBody","accent","accentBody"])),
    ];
    for (position,path) in cases {
        let result=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"selection","selection":{"kind":"range","anchor":position,"focus":position}}));
        assert_eq!(result["caret"],position);assert_eq!(result["description"]["path"],path);
        assert!(result["rectangles"].as_array().unwrap().is_empty());
    }
    let up=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"vertical","from":row(&annotation["base"]["id"],0),"down":false,"desiredX":null}));
    assert_eq!(up["position"]["nodeId"],annotation["annotation"]["id"]);
    assert_eq!(ok(&mut runtime,json!({"op":"inspectDocument","handle":1})),doc);
}
#[test]
fn operator_bounds_are_described_and_vertical_navigation_does_not_land_in_unrelated_math() {
    for (source,style) in [(r"\sum_{i=1}^{n}+z","display"),(r"\sum_{i=1}^{n}+z","text"),(r"\int_0^1+x","display")] {
        let (mut runtime,mut options)=setup(source);options["style"]=json!(style);
        let scene=layout(&mut runtime,&options,0);let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
        let parts=doc["snapshot"]["analysis"]["tree"]["entries"][0]["node"]["parts"].as_array().unwrap();
        let lower=parts.iter().find(|p|p["kind"]=="script"&&p["slot"]["script"]=="subscript").unwrap()["slot"]["argument"]["id"].clone();
        let upper=parts.iter().find(|p|p["kind"]=="script"&&p["slot"]["script"]=="superscript").unwrap()["slot"]["argument"]["id"].clone();
        for (id,role) in [(&lower,"lowerLimit"),(&upper,"upperLimit")] {
            let p=row(id,0);let r=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"selection","selection":{"kind":"range","anchor":p,"focus":p}}));
            assert_eq!(r["description"]["path"],json!(["root","largeOperator",role]));
        }
        let up=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"vertical","from":row(&lower,0),"down":false,"desiredX":null}));
        assert_eq!(up["position"]["nodeId"],upper,"{source}, {style}");
        let down=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"vertical","from":row(&upper,0),"down":true,"desiredX":null}));
        assert_eq!(down["position"]["nodeId"],lower,"{source}, {style}");
    }
}
fn collapsed(runtime:&mut Runtime,token:&Value,p:&Value)->Value {
    query(runtime,token,0,json!({"kind":"selection","selection":{"kind":"range","anchor":p,"focus":p}}))
}
fn region<'a>(scene:&'a Value,id:&Value)->&'a Value {
    scene["snapshot"]["regions"].as_array().unwrap().iter().find(|r|&r["nodeId"]==id).unwrap()
}
fn f(v:&Value)->f64 {v.as_f64().unwrap()}
#[test]
fn caret_geometry_spans_and_marks_the_slot_it_is_in() {
    let (mut runtime,options)=setup(r"\frac{a+b}{c}+x");let scene=layout(&mut runtime,&options,0);let token=&scene["sceneToken"];
    let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
    let root=&doc["snapshot"]["analysis"]["tree"];let fraction=&root["entries"][0]["node"];
    // Inside a numerator: the bar covers that slot and the slot is marked.
    let inner=collapsed(&mut runtime,token,&row(&fraction["numerator"]["id"],1));
    let numerator=&region(&scene,&fraction["numerator"]["id"])["logical"];
    let geometry=&inner["caretGeometry"];
    let (slot,bar)=(&geometry["slot"],&geometry["bar"]);
    assert_eq!((&slot["xMin"],&slot["xMax"]),(&numerator["xMin"],&numerator["xMax"]),"{inner}");
    assert_eq!(bar["xMin"],bar["xMax"]);
    assert!(f(&bar["yMin"])<=f(&numerator["yMin"])+1e-9&&f(&bar["yMax"])>=f(&numerator["yMax"])-1e-9);
    // The slot mark meets the foot of the bar.
    assert_eq!((&slot["yMin"],&slot["yMax"]),(&bar["yMin"],&bar["yMax"]));
    // In the root row the caret spans the formula's tallest content, unmarked.
    let outer=collapsed(&mut runtime,token,&row(&root["id"],1));
    let whole=&region(&scene,&fraction["id"])["logical"];
    assert!(outer["caretGeometry"]["slot"].is_null(),"{outer}");
    assert!(f(&outer["caretGeometry"]["bar"]["yMin"])<=f(&whole["yMin"])+1e-9);
    assert!(f(&outer["caretGeometry"]["bar"]["yMax"])>=f(&whole["yMax"])-1e-9);
    // A range selection paints rectangles, never a caret.
    let range=query(&mut runtime,token,0,json!({"kind":"selection","selection":{"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],2)}}));
    assert!(range["caretGeometry"].is_null());
}
#[test]
fn an_empty_slot_caret_is_not_additionally_underlined() {
    let (mut runtime,options)=setup(r"\frac{}{x}");let scene=layout(&mut runtime,&options,0);
    let doc=ok(&mut runtime,json!({"op":"inspectDocument","handle":1}));
    let fraction=&doc["snapshot"]["analysis"]["tree"]["entries"][0]["node"];
    let empty=collapsed(&mut runtime,&scene["sceneToken"],&row(&fraction["numerator"]["id"],0));
    assert!(empty["caretGeometry"]["slot"].is_null(),"{empty}");
    let filled=collapsed(&mut runtime,&scene["sceneToken"],&row(&fraction["denominator"]["id"],1));
    assert!(filled["caretGeometry"]["slot"].is_object(),"{filled}");
}
/// Every empty row of the document, as (id, logical box).
fn empty_rows(scene:&Value)->Vec<(Value,Value)> {
    scene["snapshot"]["regions"].as_array().unwrap().iter().filter(|r|r["role"]=="emptyRow")
        .map(|r|(r["nodeId"].clone(),r["logical"].clone())).collect()
}
#[test]
fn every_unfilled_slot_has_one_core_placeholder_box_and_entry_caret() {
    for (source,expected) in [(r"x^{}+y",1),(r"\frac{}{}+1",2),(r"a+\placeholder{}",1),(r"\frac{x}{y}",0)] {
        let (mut runtime,options)=setup(source);let scene=layout(&mut runtime,&options,0);
        let slots=scene["snapshot"]["emptySlots"].as_array().unwrap();
        assert_eq!(slots.len(),expected,"{source}");
        for slot in slots {
            let (rect,hit)=(&slot["rect"],&slot["hit"]);
            assert!(f(&hit["xMin"])<=f(&rect["xMin"])&&f(&hit["xMax"])>=f(&rect["xMax"])&&f(&hit["yMin"])<=f(&rect["yMin"])&&f(&hit["yMax"])>=f(&rect["yMax"]));
            for (fx,fy) in [(0.05,0.5),(0.5,0.5),(0.95,0.5),(0.5,0.1),(0.5,0.9)] {
                let x=f(&rect["xMin"])+(f(&rect["xMax"])-f(&rect["xMin"]))*fx;
                let y=f(&rect["yMin"])+(f(&rect["yMax"])-f(&rect["yMin"]))*fy;
                let hit=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"hitTest","xEm":x,"yEm":y}));
                assert_eq!(hit["position"],slot["position"],"{source}: drawn box at ({fx},{fy})");
            }
        }
    }
}
#[test]
fn a_click_anywhere_on_an_empty_slot_box_lands_in_that_slot() {
    for source in [r"x^{}+y",r"x_{}+y",r"\sum_{}^{}x",r"\frac{}{}+1",r"\sqrt[]{}",r"a^{}_{}"] {
        let (mut runtime,options)=setup(source);let scene=layout(&mut runtime,&options,0);
        let rows=empty_rows(&scene);assert!(!rows.is_empty(),"{source}");
        for (id,logical) in rows {
            let (x0,x1,y0,y1)=(f(&logical["xMin"]),f(&logical["xMax"]),f(&logical["yMin"]),f(&logical["yMax"]));
            for fx in [0.1,0.5,0.9] {
                let hit=query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"hitTest","xEm":x0+(x1-x0)*fx,"yEm":(y0+y1)/2.0}));
                assert_eq!(hit["position"],row(&id,0),"{source}: click at {fx} of empty box {id}");
            }
        }
    }
}
#[test]
fn failed_layout_does_not_retire_previous_valid_geometry() {
    let (mut runtime,options)=setup("x");let scene=layout(&mut runtime,&options,0);
    let mut bad=options.clone();bad["mathFont"]=json!("missing");
    assert_eq!(send(&mut runtime,json!({"op":"layout","handle":1,"expectedRevision":0,"options":bad,"includeSvg":false}))["status"],"error");
    query(&mut runtime,&scene["sceneToken"],0,json!({"kind":"hitTest","xEm":0,"yEm":0}));
    ok(&mut runtime,json!({"op":"destroy","handle":1,"expectedRevision":0}));
    assert_eq!(send(&mut runtime,json!({"op":"queryScene","handle":1,"expectedRevision":0,"sceneToken":scene["sceneToken"],"query":{"kind":"hitTest","xEm":0,"yEm":0}}))["error"]["code"],"invalidHandle");
}
