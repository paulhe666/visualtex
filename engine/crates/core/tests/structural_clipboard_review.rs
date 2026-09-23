//! P09 versioned structural clipboard. The custom payload is a verified,
//! identity-free semantic record bound to the LaTeX fallback; it is never a
//! deserialization shortcut around the parser or a carrier for live NodeId.
use std::collections::BTreeSet;
use serde_json::{Value,json};
use visualtex_engine_core::Runtime;

fn call(runtime:&mut Runtime,command:Value)->Value{
    serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
}
fn ok(runtime:&mut Runtime,command:Value)->Value{
    let result=call(runtime,command);assert_eq!(result["status"],"ok","{result}");result["result"].clone()
}
fn error(runtime:&mut Runtime,command:Value,code:&str)->Value{
    let result=call(runtime,command);assert_eq!(result["status"],"error","{result}");assert_eq!(result["error"]["code"],code,"{result}");result
}
fn setup(source:&str)->Runtime{let mut r=Runtime::default();ok(&mut r,json!({"op":"create","documentId":"clipboard-struct","source":source}));r}
fn editor(r:&mut Runtime)->Value{ok(r,json!({"op":"inspectEditor","handle":1}))["snapshot"].clone()}
fn action(r:&mut Runtime,revision:u32,sequence:u32,action:Value)->Value{ok(r,json!({"op":"editorAction","handle":1,"expectedRevision":revision,"eventSequence":sequence,"action":action}))["result"].clone()}
fn row(id:&Value,offset:usize)->Value{json!({"kind":"row","nodeId":id,"offset":offset,"affinity":"after"})}
fn collect_ids(value:&Value,out:&mut BTreeSet<u64>){
    match value{
        Value::Array(values)=>for value in values{collect_ids(value,out)},
        Value::Object(map)=>{if let Some(id)=map.get("id").and_then(Value::as_u64){assert!(out.insert(id));}for (key,value) in map{if key!="origin"{collect_ids(value,out);}}},
        _=>{},
    }
}
fn contains_forbidden_identity_key(value:&Value)->bool{
    match value{
        Value::Array(values)=>values.iter().any(contains_forbidden_identity_key),
        Value::Object(map)=>map.iter().any(|(key,value)|matches!(key.as_str(),"id"|"origin"|"nodeId"|"documentId"|"handle"|"sourceId")||contains_forbidden_identity_key(value)),
        _=>false,
    }
}

#[test]
fn copied_complex_formula_has_versioned_identity_free_structure_bound_to_source(){
    let source=r"\pdv{f}{x}+\left\langle a\middle|\frac{b}{c}\right\rangle+\left[\begin{array}{cc|c}a&b&c\\d&e&f\end{array}\right]";
    let mut r=setup(source);let before=editor(&mut r);let root=&before["document"]["analysis"]["tree"];
    action(&mut r,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],root["entries"].as_array().unwrap().len())}}));
    let selected=editor(&mut r);let fragment=ok(&mut r,json!({"op":"copySelection","handle":1,"expectedRevision":0}))["fragment"].clone();
    assert_eq!(fragment["schemaVersion"],1);assert_eq!(fragment["latex"],source);let structure=&fragment["structure"];
    assert_eq!(structure["schema"],"visualtex-engine-fragment");assert_eq!(structure["version"],1);assert_eq!(structure["assets"],json!([]));
    assert_eq!(structure["sourceSha256"].as_str().unwrap().len(),64);assert!(!contains_forbidden_identity_key(&structure["shape"]));
    assert_eq!(editor(&mut r),selected,"copy must remain read-only");
}

#[test]
fn structured_paste_revalidates_source_and_allocates_independent_nodes_with_history(){
    let mut r=setup(r"\frac{x}{y}");let initial=editor(&mut r);let root=&initial["document"]["analysis"]["tree"];let original=&root["entries"][0]["node"];
    action(&mut r,0,1,json!({"kind":"setSelection","selection":{"kind":"node","nodeId":original["id"]}}));
    let fragment=ok(&mut r,json!({"op":"copySelection","handle":1,"expectedRevision":0}))["fragment"].clone();let structure=fragment["structure"].clone();assert!(structure.is_object());
    let end=row(&root["id"],1);action(&mut r,0,2,json!({"kind":"setSelection","selection":{"kind":"range","anchor":end,"focus":end}}));
    let pasted=action(&mut r,0,3,json!({"kind":"edit","command":{"kind":"pasteStructured","latex":fragment["latex"],"fragment":structure}}))["editor"].clone();
    assert_eq!(pasted["document"]["source"],r"\frac{x}{y}\frac{x}{y}");assert_eq!(pasted["undoDepth"],1);
    let entries=pasted["document"]["analysis"]["tree"]["entries"].as_array().unwrap();let mut left=BTreeSet::new();let mut right=BTreeSet::new();collect_ids(&entries[0]["node"],&mut left);collect_ids(&entries[1]["node"],&mut right);assert!(left.is_disjoint(&right));
    let undone=action(&mut r,1,4,json!({"kind":"undo"}))["editor"].clone();assert_eq!(undone["document"]["source"],r"\frac{x}{y}");
    let redone=action(&mut r,2,5,json!({"kind":"redo"}))["editor"].clone();assert_eq!(redone["document"]["source"],r"\frac{x}{y}\frac{x}{y}");assert_eq!(redone["document"]["analysis"]["tree"],pasted["document"]["analysis"]["tree"]);
}

#[test]
fn tampered_digest_shape_version_and_asset_refs_fail_atomically(){
    let mut r=setup("x");let state=editor(&mut r);let root=&state["document"]["analysis"]["tree"];
    action(&mut r,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],1)}}));
    let fragment=ok(&mut r,json!({"op":"copySelection","handle":1,"expectedRevision":0}))["fragment"].clone();let good=fragment["structure"].clone();
    let end=row(&root["id"],1);action(&mut r,0,2,json!({"kind":"setSelection","selection":{"kind":"range","anchor":end,"focus":end}}));let before=editor(&mut r);
    let mut seq=3;
    let mut cases:Vec<(Value,&str)>=Vec::new();
    let mut digest=good.clone();digest["sourceSha256"]=json!("0".repeat(64));cases.push((digest,"invalidInput"));
    let mut shape=good.clone();shape["shape"]["kind"]=json!("fraction");cases.push((shape,"invalidInput"));
    let mut unknown=good.clone();unknown["shape"]["kind"]=json!("futureExecutableNode");cases.push((unknown,"invalidInput"));
    let mut version=good.clone();version["version"]=json!(99);cases.push((version,"invalidInput"));
    let mut schema=good.clone();schema["schema"]=json!("not-visualtex");cases.push((schema,"invalidInput"));
    let mut asset=good.clone();asset["assets"]=json!([{"kind":"customSymbol","assetId":"foreign","sha256":"0".repeat(64)}]);cases.push((asset,"unsupportedSyntax"));
    for (tampered,code) in cases{
        error(&mut r,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":seq,"action":{"kind":"edit","command":{"kind":"pasteStructured","latex":"x","fragment":tampered}}}),code);seq+=1;assert_eq!(editor(&mut r),before);
    }
}

#[test]
fn oversized_structure_and_unknown_wire_fields_are_rejected_without_mutation(){
    let mut r=setup("x");let state=editor(&mut r);let root=&state["document"]["analysis"]["tree"];action(&mut r,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],1)}}));let fragment=ok(&mut r,json!({"op":"copySelection","handle":1,"expectedRevision":0}))["fragment"].clone();
    let end=row(&root["id"],1);action(&mut r,0,2,json!({"kind":"setSelection","selection":{"kind":"range","anchor":end,"focus":end}}));let before=editor(&mut r);
    let mut huge=fragment["structure"].clone();huge["shape"]=json!("x".repeat(600_000));error(&mut r,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":3,"action":{"kind":"edit","command":{"kind":"pasteStructured","latex":"x","fragment":huge}}}),"resourceLimit");assert_eq!(editor(&mut r),before);
    let mut unknown=fragment["structure"].clone();unknown["unexpectedHtml"]=json!("<script>evil()</script>");error(&mut r,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":4,"action":{"kind":"edit","command":{"kind":"pasteStructured","latex":"x","fragment":unknown}}}),"invalidInput");assert_eq!(editor(&mut r),before);
}

#[test]
fn unsupported_or_incomplete_source_never_receives_a_structural_trust_marker(){
    for source in [r"\futureCommand{a}",r"\frac{x}{"]{
        let mut r=setup(source);let state=editor(&mut r);let root=&state["document"]["analysis"]["tree"];let count=root["entries"].as_array().unwrap().len();if count==0{continue;}
        action(&mut r,0,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],count)}}));
        // Draft/unsupported source cannot be visually selected through a stale retained tree in every case;
        // when copy is accepted, it must never gain a trusted structure payload.
        let reply=call(&mut r,json!({"op":"copySelection","handle":1,"expectedRevision":0}));if reply["status"]=="ok"{assert!(reply["result"]["fragment"]["structure"].is_null());}
    }
}
