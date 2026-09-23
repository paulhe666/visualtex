//! Independent single-writer, revision and immutable-snapshot contract checks.
//! Full-source replacement intentionally creates fresh inner identities; stable
//! local edits and undo/redo belong to P04 and are not inferred from these tests.
use std::collections::BTreeSet;
use std::sync::Arc;
use serde_json::{Value, json};
use visualtex_engine_core::{ErrorCode, Runtime};
use visualtex_engine_core::model::{LatexMode, Node, NodeKind, latex};
use visualtex_engine_core::session::{EngineSession, Revision};

fn snapshot(session: &EngineSession) -> Value {
    serde_json::to_value(session.document_snapshot()).unwrap()
}
fn inner_ids(root: &Node) -> BTreeSet<u32> {
    let mut ids = BTreeSet::new();
    root.visit(&mut |node| { if node.id != root.id { ids.insert(node.id.0); } });
    ids
}
fn send(runtime: &mut Runtime, command: Value) -> Value {
    serde_json::from_str(&runtime.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
}

#[test]
fn noops_preserve_revision_source_identity_and_every_node() {
    let mut session = EngineSession::new("document".into()).unwrap();
    for source in ["", "x", "\\frac{x_1}{y^2}", "\\text{速度 👩‍🔬}", "\\frac{a}{", "\\unknown[a]{b}"] {
        let changed = session.replace_source(session.revision(), source.into()).unwrap();
        let before = serde_json::to_value(&changed).unwrap();
        for _ in 0..5 {
            assert_eq!(serde_json::to_value(session.replace_source(session.revision(), source.into()).unwrap()).unwrap(), before);
        }
    }
}

#[test]
fn stale_revision_is_rejected_even_for_noop_and_overbudget_input() {
    let mut session = EngineSession::new("document".into()).unwrap();
    session.replace_source(Revision(0), "x".into()).unwrap();
    let before = snapshot(&session);
    for source in ["x".to_owned(), "y".to_owned(), "x".repeat(300_000)] {
        assert_eq!(session.replace_source(Revision(0), source).unwrap_err().code, ErrorCode::StaleRevision);
        assert_eq!(snapshot(&session), before);
    }
}

#[test]
fn failed_parse_budgets_do_not_consume_future_node_or_source_identities() {
    let mut subject = EngineSession::new("document".into()).unwrap();
    let mut control = EngineSession::new("document".into()).unwrap();
    let seed = "\\frac{x_1}{y^2}";
    subject.replace_source(Revision(0), seed.into()).unwrap();
    control.replace_source(Revision(0), seed.into()).unwrap();
    for bad in ["x".repeat(9000), "\\unknown ".repeat(200),
        format!("{}x{}", "{".repeat(100), "}".repeat(100)), "x".repeat(262145)] {
        let before = snapshot(&subject);
        assert_eq!(subject.replace_source(subject.revision(), bad).unwrap_err().code, ErrorCode::ResourceLimit);
        assert_eq!(snapshot(&subject), before);
    }
    let after_subject = subject.replace_source(subject.revision(), "\\frac{a+1}{b_2}".into()).unwrap();
    let after_control = control.replace_source(control.revision(), "\\frac{a+1}{b_2}".into()).unwrap();
    assert_eq!(serde_json::to_value(after_subject).unwrap(), serde_json::to_value(after_control).unwrap());
}

#[test]
fn explicit_full_replacement_never_reuses_inner_identities() {
    let mut session = EngineSession::new("document".into()).unwrap();
    let mut all_inner_ids = BTreeSet::new();
    for source in ["x+x", "y+y", "x+x", "\\frac{x}{x}", "x+x"] {
        let state = session.replace_source(session.revision(), source.into()).unwrap();
        let current = inner_ids(&state.analysis.tree);
        assert!(all_inner_ids.is_disjoint(&current));
        all_inner_ids.extend(current);
        assert_eq!(state.analysis.tree.id.0, 1);
        assert_eq!(latex(&state.analysis.tree, LatexMode::Lossless), source);
    }
}

#[test]
fn exported_snapshot_is_copy_on_write_and_cannot_mutate_owner() {
    let mut session = EngineSession::new("document".into()).unwrap();
    session.replace_source(Revision(0), "x+1".into()).unwrap();
    let before = snapshot(&session);
    let mut external = session.document_snapshot();
    external.source = "changed only in snapshot".into();
    let analysis = Arc::make_mut(&mut external.analysis);
    let tree = Arc::make_mut(&mut analysis.tree);
    tree.origin = None;
    tree.kind = NodeKind::UnparsedSource { raw: "external".into() };
    assert_eq!(snapshot(&session), before);
}

#[test]
fn draft_cycles_retain_old_tree_with_explicit_old_source_identity() {
    let mut session = EngineSession::new("document".into()).unwrap();
    let good = session.replace_source(Revision(0), "\\frac{x_1}{y^2}".into()).unwrap();
    let good_source_id = good.source_id;
    let good_tree = serde_json::to_value(&good.analysis.tree).unwrap();
    for draft in ["\\frac{x}{", "\\frac{x}", "\\text{速度", "\\"] {
        let state = session.replace_source(session.revision(), draft.into()).unwrap();
        assert_eq!(state.source, draft);
        assert!(!state.analysis.complete);
        assert!(state.canonical_latex.is_none());
        let retained = state.retained_tree.as_ref().expect("previous complete tree is explicitly retained");
        assert_eq!(serde_json::to_value(retained).unwrap(), good_tree);
        assert_eq!(retained.origin.as_ref().unwrap().span.source_id, good_source_id);
        assert_ne!(good_source_id, state.source_id);
        assert_eq!(latex(&state.analysis.tree, LatexMode::Lossless), draft);
    }
    let recovered = session.replace_source(session.revision(), "\\frac{x}{2}".into()).unwrap();
    assert!(recovered.analysis.exportable);
    assert!(recovered.retained_tree.is_none());
    assert_eq!(recovered.canonical_latex.as_deref(), Some("\\frac{x}{2}"));
}

#[test]
fn source_protocol_requires_handle_and_revision_and_is_document_isolated() {
    let mut runtime = Runtime::default();
    for id in ["doc-a", "doc-b"] {
        assert_eq!(send(&mut runtime, json!({"op":"create","documentId":id}))["status"], "ok");
    }
    let before_b = send(&mut runtime, json!({"op":"inspectDocument","handle":2}));
    let first = send(&mut runtime, json!({"op":"replaceSource","handle":1,"expectedRevision":0,"source":"x_1"}));
    assert_eq!(first["status"], "ok");
    for command in [json!({"op":"replaceSource","handle":1,"source":"bad"}),
        json!({"op":"replaceSource","handle":1,"expectedRevision":1,"source":"bad","bogus":1}),
        json!({"op":"replaceSource","handle":1,"expectedRevision":1,"source":false})] {
        assert_eq!(send(&mut runtime, command)["error"]["code"], "invalidInput");
    }
    assert_eq!(send(&mut runtime, json!({"op":"replaceSource","handle":1,"expectedRevision":0,"source":"wrong"}))["error"]["code"], "staleRevision");
    assert_eq!(send(&mut runtime, json!({"op":"replaceSource","handle":77,"expectedRevision":0,"source":"wrong"}))["error"]["code"], "invalidHandle");
    assert_eq!(send(&mut runtime, json!({"op":"inspectDocument","handle":2})), before_b);
    let a = send(&mut runtime, json!({"op":"inspect","handle":1}));
    assert_eq!(a["result"]["snapshot"]["source"], "x_1");
    assert_eq!(a["result"]["snapshot"]["revision"], 1);
    assert_eq!(runtime.active_sessions(), 2);
}

#[test]
fn old_snapshot_remains_readable_after_owner_replacement_and_disposal() {
    let old;
    {
        let mut session = EngineSession::new("document".into()).unwrap();
        old = session.replace_source(Revision(0), "\\frac{速度}{😀}".into()).unwrap();
        session.replace_source(Revision(1), "z".into()).unwrap();
    }
    assert_eq!(old.source, "\\frac{速度}{😀}");
    assert_eq!(latex(&old.analysis.tree, LatexMode::Lossless), old.source);
    assert_eq!(old.analysis.source.text(), old.source);
}
