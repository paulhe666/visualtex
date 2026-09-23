use serde_json::{Value,json};
use std::collections::BTreeSet;
use visualtex_engine_core::Runtime;
fn call(runtime:&mut Runtime,command:Value)->Value {
    serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
}
fn ok(runtime:&mut Runtime,command:Value)->Value {
    let result=call(runtime,command);assert_eq!(result["status"],"ok","{result}");result["result"].clone()
}
fn setup(source:&str)->Runtime {
    let mut runtime=Runtime::default();ok(&mut runtime,json!({"op":"create","documentId":"clipboard","source":source}));runtime
}
fn editor(runtime:&mut Runtime)->Value {ok(runtime,json!({"op":"inspectEditor","handle":1}))["snapshot"].clone()}
fn action(runtime:&mut Runtime,revision:u32,sequence:u32,action:Value)->Value {
    ok(runtime,json!({"op":"editorAction","handle":1,"expectedRevision":revision,"eventSequence":sequence,"action":action}))["result"].clone()
}
fn copy(runtime:&mut Runtime,revision:u32)->Value {ok(runtime,json!({"op":"copySelection","handle":1,"expectedRevision":revision}))["fragment"].clone()}
fn row(id:&Value,offset:usize)->Value {json!({"kind":"row","nodeId":id,"offset":offset,"affinity":"after"})}
fn text(id:&Value,grapheme:usize)->Value {json!({"kind":"text","nodeId":id,"grapheme":grapheme,"affinity":"after"})}
fn ids(node:&Value,output:&mut BTreeSet<u64>) {
    if let Some(id)=node.get("id").and_then(Value::as_u64) {assert!(output.insert(id));}
    match node["kind"].as_str().unwrap() {
        "row"=>for entry in node["entries"].as_array().unwrap(){ids(&entry["node"],output)},
        "fraction"=>{ids(&node["numerator"],output);ids(&node["denominator"],output);},
        "scripts"=>{ids(&node["base"],output);for slot in node["slots"].as_array().unwrap(){ids(&slot["argument"],output);}},_=>{},
    }
}
#[test]
fn copying_a_formula_is_read_only_and_paste_allocates_independent_nodes() {
    let mut runtime=setup(r"\frac{x}{y}+1");let initial=editor(&mut runtime);
    let root=&initial["document"]["analysis"]["tree"];let original=&root["entries"][0]["node"];
    action(&mut runtime,0,1,json!({"kind":"setSelection","selection":{"kind":"node","nodeId":original["id"]}}));
    let selected=editor(&mut runtime);let fragment=copy(&mut runtime,0);
    assert_eq!(fragment["latex"],r"\frac{x}{y}");assert_eq!(fragment["schemaVersion"],1);
    assert_eq!(editor(&mut runtime),selected);assert!(!fragment.to_string().contains("nodeId"));
    let end=row(&root["id"],3);
    action(&mut runtime,0,2,json!({"kind":"setSelection","selection":{"kind":"range","anchor":end,"focus":end}}));
    let result=action(&mut runtime,0,3,json!({"kind":"edit","command":{"kind":"pasteLatex","latex":fragment["latex"]}}));
    let pasted=&result["editor"]["document"]["analysis"]["tree"]["entries"][3]["node"];
    let mut original_ids=BTreeSet::new();ids(original,&mut original_ids);
    let mut pasted_ids=BTreeSet::new();ids(pasted,&mut pasted_ids);assert!(original_ids.is_disjoint(&pasted_ids));
    assert_eq!(result["editor"]["document"]["source"],r"\frac{x}{y}+1\frac{x}{y}");
}
#[test]
fn partial_text_copy_preserves_whole_graphemes_and_escapes_source_metacharacters() {
    let mut runtime=setup("\\text{A😀e\u{301}\\{B}");let state=editor(&mut runtime);
    let id=&state["document"]["analysis"]["tree"]["entries"][0]["node"]["id"];
    action(&mut runtime,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":text(id,1),"focus":text(id,4)}}));
    let fragment=copy(&mut runtime,0);assert_eq!(fragment["latex"],"\\text{😀e\u{301}\\{}");
    let p=text(id,5);
    action(&mut runtime,0,2,json!({"kind":"setSelection","selection":{"kind":"range","anchor":p,"focus":p}}));
    let result=action(&mut runtime,0,3,json!({"kind":"edit","command":{"kind":"pasteLatex","latex":fragment["latex"]}}));
    assert_eq!(result["editor"]["document"]["analysis"]["complete"],true);
    let text_values:Vec<_>=result["editor"]["document"]["analysis"]["tree"]["entries"].as_array().unwrap().iter().filter_map(|e|e["node"]["text"].as_str()).collect();
    assert_eq!(text_values.concat(),"A😀e\u{301}{B😀e\u{301}{");
}
#[test]
fn no_selection_does_not_replace_the_system_clipboard_with_empty_text() {
    let mut runtime=setup("x");assert!(copy(&mut runtime,0)["latex"].is_null());
}
#[test]
fn stale_revisions_and_active_compositions_cannot_be_cut_as_committed_text() {
    let mut runtime=setup(r"\text{AB}");
    assert_eq!(call(&mut runtime,json!({"op":"copySelection","handle":1,"expectedRevision":1}))["error"]["code"],"staleRevision");
    action(&mut runtime,0,1,json!({"kind":"beginComposition","mode":"text"}));
    assert_eq!(call(&mut runtime,json!({"op":"copySelection","handle":1,"expectedRevision":0}))["error"]["code"],"conflict");
    assert_eq!(editor(&mut runtime)["document"]["source"],r"\text{AB}");
}
#[test]
fn row_copy_retains_control_word_boundaries_without_borrowing_outer_trivia() {
    let mut runtime=setup("  \\alpha x");let state=editor(&mut runtime);let id=&state["document"]["analysis"]["tree"]["id"];
    action(&mut runtime,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":row(id,0),"focus":row(id,2)}}));
    let copied=copy(&mut runtime,0);assert_eq!(copied["latex"],r"\alpha x");
    let parsed=ok(&mut runtime,json!({"op":"analyzeSource","source":copied["latex"]}));
    assert_eq!(parsed["analysis"]["supported"],true);assert_eq!(parsed["analysis"]["facts"]["nodes"],3);
}
