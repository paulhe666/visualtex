//! P08 long-document retention: parked formula sessions preserve the complete
//! editor owner while releasing the active layout/input budget.
use serde_json::{json,Value};
use visualtex_engine_core::Runtime;

fn send(runtime:&mut Runtime,command:Value)->Value {
    serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
}
fn editor(runtime:&mut Runtime,handle:u32)->Value {
    send(runtime,json!({"op":"inspectEditor","handle":handle}))["result"]["snapshot"].clone()
}

#[test]
fn one_thousand_formula_sessions_park_without_losing_identity_history_or_event_receipts() {
    let mut runtime=Runtime::default();
    let sentinels=[1u32,501,1000];
    let mut saved=Vec::new();
    for index in 0..1000u32 {
        let handle=index+1;
        let created=send(&mut runtime,json!({"op":"create","documentId":format!("long-formula-{index}"),"source":"x"}));
        assert_eq!(created["status"],"ok");assert_eq!(created["result"]["handle"],handle);
        let digit=char::from(b'0'+(index%10)as u8).to_string();
        let action=json!({"kind":"edit","command":{"kind":"insertText","text":digit,"mode":"math"},"groupId":null});
        let edited=send(&mut runtime,json!({"op":"editorAction","handle":handle,"expectedRevision":0,"eventSequence":1,"action":action}));
        assert_eq!(edited["status"],"ok");
        if sentinels.contains(&handle){saved.push((handle,action,edited["result"]["result"]["editor"].clone()));}
        let parked=send(&mut runtime,json!({"op":"parkSession","handle":handle,"expectedRevision":1}));
        assert_eq!(parked["status"],"ok");assert_eq!(parked["result"]["handle"],handle);
        assert_eq!(runtime.active_sessions(),0,"Parked formula consumed an active slot at {handle}");
        assert_eq!(runtime.parked_sessions(),handle as usize);
    }
    assert_eq!(runtime.active_sessions(),0);assert_eq!(runtime.parked_sessions(),1000);
    assert!(runtime.parked_weight()>0);

    for (handle,first_action,before) in &saved {
        let resumed=send(&mut runtime,json!({"op":"resumeSession","handle":handle,"expectedRevision":1}));
        assert_eq!(resumed["status"],"ok");assert_eq!(resumed["result"]["handle"],*handle);
        assert_eq!(editor(&mut runtime,*handle),*before,"Park/resume changed the complete editor snapshot");
        // Event receipts live in the EngineSession too: the original event is
        // acknowledged as a replay rather than inserted a second time.
        let replay=send(&mut runtime,json!({"op":"editorAction","handle":handle,"expectedRevision":0,
            "eventSequence":1,"action":first_action}));
        assert_eq!(replay["status"],"ok");assert_eq!(replay["result"]["result"]["replayed"],true);
        assert_eq!(replay["result"]["result"]["editor"]["document"]["source"],before["document"]["source"]);
        let undo=send(&mut runtime,json!({"op":"editorAction","handle":handle,"expectedRevision":1,"eventSequence":2,"action":{"kind":"undo"}}));
        assert_eq!(undo["result"]["result"]["editor"]["document"]["source"],"x");
        let redo=send(&mut runtime,json!({"op":"editorAction","handle":handle,"expectedRevision":2,"eventSequence":3,"action":{"kind":"redo"}}));
        let after=&redo["result"]["result"]["editor"];
        assert_eq!(after["document"]["source"],before["document"]["source"]);
        assert_eq!(after["document"]["analysis"]["tree"],before["document"]["analysis"]["tree"],"Node identities changed after parked history replay");
        assert_eq!(after["selection"],before["selection"]);
        assert_eq!(send(&mut runtime,json!({"op":"parkSession","handle":handle,"expectedRevision":3}))["status"],"ok");
    }
    assert_eq!(runtime.active_sessions(),0);assert_eq!(runtime.parked_sessions(),1000);

    for handle in 1..=1000u32 {
        let revision=if sentinels.contains(&handle){3}else{1};
        assert_eq!(send(&mut runtime,json!({"op":"destroy","handle":handle,"expectedRevision":revision}))["status"],"ok");
    }
    assert_eq!(runtime.active_sessions(),0);assert_eq!(runtime.parked_sessions(),0);assert_eq!(runtime.parked_weight(),0);
}

#[test]
fn active_budget_remains_bounded_while_parked_documents_keep_their_ids() {
    let mut runtime=Runtime::default();
    for index in 0..128u32 {
        let handle=index+1;
        assert_eq!(send(&mut runtime,json!({"op":"create","documentId":format!("bounded-{index}"),"source":"x+y"}))["status"],"ok");
        assert_eq!(send(&mut runtime,json!({"op":"parkSession","handle":handle,"expectedRevision":0}))["status"],"ok");
    }
    assert_eq!(runtime.active_sessions(),0);assert_eq!(runtime.parked_sessions(),128);
    // Reopening a retained formula never allocates a replacement handle.
    for handle in [1u32,64,128] {
        assert_eq!(send(&mut runtime,json!({"op":"resumeSession","handle":handle,"expectedRevision":0}))["result"]["handle"],handle);
        assert_eq!(send(&mut runtime,json!({"op":"parkSession","handle":handle,"expectedRevision":0}))["result"]["handle"],handle);
    }
}
