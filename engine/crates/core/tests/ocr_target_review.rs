//! Deferred OCR insertion uses an owned, one-shot target, never the current host.
use serde_json::{json,Value};
use visualtex_engine_core::Runtime;
fn send(rt:&mut Runtime,c:Value)->Value{serde_json::from_str(&rt.dispatch_json(&json!({"protocol":1,"command":c}).to_string())).unwrap()}
fn ok(rt:&mut Runtime,c:Value)->Value{let r=send(rt,c);assert_eq!(r["status"],"ok","{r}");r["result"].clone()}
fn create(rt:&mut Runtime,name:&str,source:&str)->u64{ok(rt,json!({"op":"create","documentId":name,"source":source}))["handle"].as_u64().unwrap()}
fn inspect(rt:&mut Runtime,h:u64)->Value{ok(rt,json!({"op":"inspectEditor","handle":h}))["snapshot"].clone()}
fn capture(rt:&mut Runtime,h:u64)->Value{let rev=inspect(rt,h)["document"]["revision"].clone();ok(rt,json!({"op":"captureOcrTarget","handle":h,"expectedRevision":rev}))["target"].clone()}
fn action(rt:&mut Runtime,h:u64,seq:u32,a:Value)->Value{let rev=inspect(rt,h)["document"]["revision"].clone();ok(rt,json!({"op":"editorAction","handle":h,"expectedRevision":rev,"eventSequence":seq,"action":a}))["result"].clone()}
fn request(h:u64,seq:u32,t:&Value,latex:&str)->Value{json!({"op":"editorAction","handle":h,"expectedRevision":t["revision"],"eventSequence":seq,"action":{"kind":"applyOcrLatex","target":t,"latex":latex}})}
fn reject(rt:&mut Runtime,h:u64,c:Value,code:&str){let before=inspect(rt,h);let r=send(rt,c);assert_eq!(r["status"],"error","{r}");assert_eq!(r["error"]["code"],code,"{r}");assert_eq!(inspect(rt,h),before,"rejection mutated source/selection/history");}
#[test]
fn ocr_capture_is_read_only_and_acceptance_is_one_undoable_transaction(){
 let mut rt=Runtime::default();let h=create(&mut rt,"ocr","x");let before=inspect(&mut rt,h);let t=capture(&mut rt,h);assert_eq!(inspect(&mut rt,h),before);assert_eq!(t["sessionHandle"],h);
 let c=request(h,1,&t,r"\frac{a}{b}");let r=ok(&mut rt,c.clone());assert_eq!(r["result"]["editor"]["document"]["source"],r"x\frac{a}{b}");assert_eq!(r["result"]["editor"]["undoDepth"],1);
 let done=inspect(&mut rt,h);assert_eq!(ok(&mut rt,c)["result"]["replayed"],true);assert_eq!(inspect(&mut rt,h),done);
 action(&mut rt,h,2,json!({"kind":"undo"}));let undone=inspect(&mut rt,h);assert_eq!(undone["document"]["source"],before["document"]["source"]);assert_eq!(undone["selection"],before["selection"]);
 action(&mut rt,h,3,json!({"kind":"redo"}));assert_eq!(inspect(&mut rt,h)["document"]["source"],done["document"]["source"]);
}
#[test]
fn selection_away_and_back_does_not_revive_ocr_target(){
 let mut rt=Runtime::default();let h=create(&mut rt,"selection","xy");let t=capture(&mut rt,h);action(&mut rt,h,1,json!({"kind":"moveHorizontal","direction":"left","extend":false}));action(&mut rt,h,2,json!({"kind":"setSelection","selection":t["selection"]}));
 assert_eq!(inspect(&mut rt,h)["document"]["revision"],t["revision"]);reject(&mut rt,h,request(h,3,&t,"z"),"conflict");
}
#[test]
fn source_edit_then_undo_cannot_revive_old_ocr_revision(){
 let mut rt=Runtime::default();let h=create(&mut rt,"revision","x");let t=capture(&mut rt,h);action(&mut rt,h,1,json!({"kind":"edit","command":{"kind":"insertText","text":"y","mode":"math"},"groupId":null}));action(&mut rt,h,2,json!({"kind":"undo"}));assert_eq!(inspect(&mut rt,h)["document"]["source"],"x");reject(&mut rt,h,request(h,3,&t,"z"),"staleRevision");
}
#[test]
fn duplicate_names_are_rejected_and_ocr_request_ownership_remains_isolated(){
 let mut rt=Runtime::default();let a=create(&mut rt,"owner-a","x");assert_eq!(send(&mut rt,json!({"op":"create","documentId":"owner-a","source":"x"}))["error"]["code"],"conflict");let b=create(&mut rt,"owner-b","x");let t=capture(&mut rt,a);let tb=capture(&mut rt,b);let before_a=inspect(&mut rt,a);reject(&mut rt,b,request(b,1,&t,"q"),"conflict");assert_eq!(inspect(&mut rt,a),before_a);ok(&mut rt,request(b,1,&tb,"b"));ok(&mut rt,request(a,1,&t,"a"));assert_eq!(inspect(&mut rt,a)["document"]["source"],"xa");assert_eq!(inspect(&mut rt,b)["document"]["source"],"xb");
}
#[test]
fn newer_ocr_request_supersedes_old_and_cancel_is_scoped(){
 let mut rt=Runtime::default();let h=create(&mut rt,"supersede","x");let a=capture(&mut rt,h);let b=capture(&mut rt,h);assert_ne!(a["requestId"],b["requestId"]);reject(&mut rt,h,request(h,1,&a,"a"),"conflict");assert_eq!(ok(&mut rt,json!({"op":"cancelOcrTarget","handle":h,"requestId":a["requestId"]}))["cancelled"],false);
 assert_eq!(ok(&mut rt,json!({"op":"cancelOcrTarget","handle":h,"requestId":b["requestId"]}))["cancelled"],true);reject(&mut rt,h,request(h,1,&b,"b"),"conflict");
}
#[test]
fn ime_begin_and_cancel_permanently_invalidates_prior_ocr_request(){
 let mut rt=Runtime::default();let h=create(&mut rt,"ime","x");let t=capture(&mut rt,h);let started=action(&mut rt,h,1,json!({"kind":"beginComposition","mode":"math"}));let id=started["editor"]["composition"]["id"].clone();reject(&mut rt,h,json!({"op":"captureOcrTarget","handle":h,"expectedRevision":0}),"conflict");reject(&mut rt,h,request(h,2,&t,"y"),"conflict");action(&mut rt,h,2,json!({"kind":"cancelComposition","id":id}));reject(&mut rt,h,request(h,3,&t,"y"),"conflict");
}
#[test]
fn parked_destroyed_and_reopened_sessions_do_not_reuse_ocr_targets(){
 let mut rt=Runtime::default();let h=create(&mut rt,"lifecycle","x");let t=capture(&mut rt,h);ok(&mut rt,json!({"op":"parkSession","handle":h,"expectedRevision":0}));ok(&mut rt,json!({"op":"resumeSession","handle":h,"expectedRevision":0}));reject(&mut rt,h,request(h,1,&t,"z"),"conflict");ok(&mut rt,json!({"op":"destroy","handle":h,"expectedRevision":0}));let next=create(&mut rt,"lifecycle","x");assert_ne!(next,h);reject(&mut rt,next,request(next,1,&t,"z"),"conflict");assert_eq!(send(&mut rt,request(h,1,&t,"z"))["error"]["code"],"invalidHandle");
}
#[test]
fn malformed_unknown_empty_or_tampered_ocr_results_are_atomic(){
 let mut rt=Runtime::default();let h=create(&mut rt,"bad-result","x");let t=capture(&mut rt,h);
 for latex in ["", " \n "]{reject(&mut rt,h,request(h,1,&t,latex),"invalidInput");}
 for (field,value) in [("documentId",json!("wrong")),("sourceId",json!(9999)),("selectionEpoch",json!(9999)),("requestId",json!(9999))]{let mut forged=t.clone();forged[field]=value;reject(&mut rt,h,request(h,1,&forged,"q"),"conflict");}
 for latex in [r"\frac{",r"\unknownocrcommand{x}"]{let before=inspect(&mut rt,h);let r=send(&mut rt,request(h,1,&t,latex));assert_eq!(r["status"],"error","{r}");assert_eq!(inspect(&mut rt,h),before);}
 ok(&mut rt,request(h,1,&t,"q"));assert_eq!(inspect(&mut rt,h)["document"]["source"],"xq");
}
#[test]
fn unfinished_and_unsupported_sources_cannot_capture_visual_ocr_targets(){
 let mut rt=Runtime::default();for (source,code) in [(r"\frac{","incompleteSource"),(r"\unknownocr{x}","unsupportedSyntax")]{let h=create(&mut rt,code,source);reject(&mut rt,h,json!({"op":"captureOcrTarget","handle":h,"expectedRevision":0}),code);}
}
#[test]
fn ocr_replaces_original_math_selection_with_fresh_structure_and_one_history_item(){
 let mut rt=Runtime::default();let h=create(&mut rt,"range","x+y");let tree=inspect(&mut rt,h)["document"]["analysis"]["tree"].clone();let root=tree["id"].clone();let p=|offset|json!({"kind":"row","nodeId":root,"offset":offset,"affinity":"after"});action(&mut rt,h,1,json!({"kind":"setSelection","selection":{"kind":"range","anchor":p(0),"focus":p(3)}}));let t=capture(&mut rt,h);ok(&mut rt,request(h,2,&t,r"\sqrt{z}"));let s=inspect(&mut rt,h);assert_eq!(s["document"]["source"],r"\sqrt{z}");assert_eq!(s["undoDepth"],1);action(&mut rt,h,3,json!({"kind":"undo"}));assert_eq!(inspect(&mut rt,h)["document"]["source"],"x+y");
}
