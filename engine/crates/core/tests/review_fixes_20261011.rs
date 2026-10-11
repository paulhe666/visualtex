//! Regression tests for the 2026-10-11 kernel/caret review fixes.
use std::path::PathBuf;
use serde_json::{Value,json};
use visualtex_engine_core::Runtime;

struct H { rt: Runtime, rev: u32, seq: u32, opts: Value }
impl H {
    fn new(src:&str)->H {
        let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
        let mut rt=Runtime::default();
        let font=rt.register_font(&std::fs::read(path).unwrap(),0).unwrap();
        let mut h=H{rt,rev:0,seq:0,opts:json!({"mathFont":font.id,"textFont":null,"fontSizePx":16,"paddingEm":0.125,"style":"display"})};
        h.ok(json!({"op":"create","documentId":"review","source":src})); h
    }
    fn send(&mut self,c:Value)->Value{serde_json::from_str(&self.rt.dispatch_json(&json!({"protocol":1,"command":c}).to_string())).unwrap()}
    fn ok(&mut self,c:Value)->Value{let r=self.send(c);assert_eq!(r["status"],"ok","{r}");r["result"].clone()}
    fn act(&mut self,a:Value)->Value{
        self.seq+=1;
        let r=self.send(json!({"op":"editorAction","handle":1,"expectedRevision":self.rev,"eventSequence":self.seq,"action":a}));
        if r["status"]=="ok" { self.rev=self.editor()["document"]["revision"].as_u64().unwrap() as u32; }
        r
    }
    fn act_ok(&mut self,a:Value)->Value{let r=self.act(a);assert_eq!(r["status"],"ok","{r}");r}
    fn edit(&mut self,cmd:Value)->Value{self.act_ok(json!({"kind":"edit","command":cmd,"groupId":null}))}
    fn editor(&mut self)->Value{self.ok(json!({"op":"inspectEditor","handle":1}))["snapshot"].clone()}
    fn src(&mut self)->String{self.editor()["document"]["source"].as_str().unwrap().to_owned()}
    fn tree(&mut self)->Value{self.editor()["document"]["analysis"]["tree"].clone()}
    fn sel(&mut self)->Value{self.editor()["selection"].clone()}
    fn layout(&mut self)->Value{let o=self.opts.clone();let rev=self.rev;self.ok(json!({"op":"layout","handle":1,"expectedRevision":rev,"options":o,"includeSvg":false}))}
    fn query(&mut self,q:Value)->Value{let l=self.layout();let rev=self.rev;self.ok(json!({"op":"queryScene","handle":1,"expectedRevision":rev,"sceneToken":l["sceneToken"],"query":q}))["result"].clone()}
    fn set(&mut self,s:Value){self.act_ok(json!({"kind":"setSelection","selection":s}));}
    fn key(&mut self,dir:&str,extend:bool){self.act_ok(json!({"kind":"moveHorizontal","direction":dir,"extend":extend}));}
}
fn row(id:&Value,o:u32)->Value{json!({"kind":"row","nodeId":id,"offset":o,"affinity":"after"})}
fn caret(p:Value)->Value{json!({"kind":"range","anchor":p,"focus":p})}

#[test]
fn common_latex_spellings_are_supported_lossless_laid_out_and_exported() {
    for source in [r"f'(x)", r"y''", r"f'^{2}", r"a~b", r"a\ b", r"\|v\|", r"a\not=b", r"x\not\in A", r"a\bmod b",
        r"a\equiv b\pmod{p}", r"\deg f", r"1,\dots,n", r"\big(x\big)", r"\Bigl[x\Bigr]", r"\bigg|", r"\Bigm|",
        r"\displaystyle\sum_{i}x_i", r"\frac{\textstyle a}{b}", r"A\iff B", r"A\implies B", r"A\impliedby B",
        r"A\xrightarrow{f}B", r"A\xleftarrow{n\to\infty}B", r"\cfrac{1}{1+\cfrac{1}{x}}", r"\sum_{\substack{i<n\\j<m}}a_{ij}"] {
        let mut h=H::new(source);
        let analysis=h.editor()["document"]["analysis"].clone();
        assert_eq!(analysis["complete"],true,"{source}: {analysis}");
        assert_eq!(analysis["supported"],true,"{source}: {}",analysis["diagnostics"]);
        assert_eq!(h.src(),source);
        let layout=h.layout();
        assert!(layout["snapshot"]["glyphs"].as_array().unwrap().len()>0,"{source}");
        let rev=h.rev;
        let mathml=h.ok(json!({"op":"exportPresentationMathMl","handle":1,"expectedRevision":rev,"displayMode":true}));
        assert!(mathml["result"]["mathMl"].as_str().unwrap().starts_with("<math"),"{source}");
    }
}

#[test]
fn unsupported_variants_of_new_spellings_stay_fail_closed() {
    for source in [r"\xrightarrow[a]{b}", r"\not\alpha", r"\big.", r"\substack{a&b}"] {
        let mut h=H::new(source);
        assert_eq!(h.editor()["document"]["analysis"]["supported"],false,"{source}");
        assert_eq!(h.src(),source);
    }
}

#[test]
fn prime_is_drawn_like_a_superscript_prime() {
    let mut a=H::new("f'"); let mut b=H::new(r"f^{\prime}");
    let ga=a.layout()["snapshot"]["glyphs"].clone(); let gb=b.layout()["snapshot"]["glyphs"].clone();
    let prime=|g:&Value|g.as_array().unwrap()[1].clone();
    assert!((prime(&ga)["y"].as_f64().unwrap()-prime(&gb)["y"].as_f64().unwrap()).abs()<1e-6,"{ga} {gb}");
    assert!((prime(&ga)["scale"].as_f64().unwrap()-prime(&gb)["scale"].as_f64().unwrap()).abs()<1e-6);
}

#[test]
fn display_style_declaration_enlarges_following_siblings_only() {
    let mut a=H::new(r"\frac{\sum x}{2}"); let mut b=H::new(r"\frac{\displaystyle\sum x}{2}");
    let ha=a.layout()["snapshot"]["heightEm"].as_f64().unwrap(); let hb=b.layout()["snapshot"]["heightEm"].as_f64().unwrap();
    assert!(hb>ha+0.1,"{ha} {hb}");
}

#[test]
fn fixed_size_delimiters_grow_with_their_size_command() {
    let height=|s:&str|{let mut h=H::new(s);h.layout()["snapshot"]["heightEm"].as_f64().unwrap()};
    let (a,b,c)=(height(r"\big("),height(r"\Big("),height(r"\Bigg("));
    assert!(a<b&&b<c,"{a} {b} {c}");
}

#[test]
fn typing_prime_and_tie_inserts_the_same_atoms_as_their_source() {
    let mut h=H::new("f");
    h.edit(json!({"kind":"insertText","text":"'","mode":"math"}));
    h.edit(json!({"kind":"insertText","text":"~","mode":"math"}));
    h.edit(json!({"kind":"insertText","text":"x","mode":"math"}));
    assert_eq!(h.src(),"f'~x");
    let reopened=H::new("f'~x").tree();
    assert_eq!(h.tree()["entries"][1]["node"]["command"],reopened["entries"][1]["node"]["command"]);
}

#[test]
fn backspace_in_an_empty_script_removes_the_script_not_the_base() {
    let mut h=H::new("x");
    h.edit(json!({"kind":"insertStructure","structure":"superscript"}));
    assert_eq!(h.src(),"x^{}");
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"x");
    assert_eq!(h.sel(),caret(row(&json!(1),1)));
    h.edit(json!({"kind":"insertText","text":"y","mode":"math"}));
    assert_eq!(h.src(),"xy");
}

#[test]
fn backspace_in_an_empty_fraction_removes_it_and_keeps_the_neighbor() {
    let mut h=H::new("a");
    h.edit(json!({"kind":"insertStructure","structure":"fraction"}));
    assert_eq!(h.src(),"a\\frac{}{}");
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"a");
    assert_eq!(h.sel(),caret(row(&json!(1),1)));
    h.act_ok(json!({"kind":"undo"}));
    assert_eq!(h.src(),"a\\frac{}{}");
}

#[test]
fn empty_slot_removal_covers_bounds_radicals_and_remaining_scripts() {
    // x^{2}_{} with the caret in the empty subscript keeps the superscript.
    let mut h=H::new("x^{2}_{}");
    let sub=h.tree()["entries"][0]["node"]["slots"][1]["argument"]["id"].clone();
    h.set(caret(row(&sub,0)));
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"x^{2}");
    // An empty lower bound of a large operator is dropped alone.
    let mut h=H::new("\\sum_{}^{n}");
    let lower=h.tree()["entries"][0]["node"]["parts"][0]["slot"]["argument"]["id"].clone();
    h.set(caret(row(&lower,0)));
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"\\sum^{n}");
    // Delete (forward) inside an empty radical removes it too.
    let mut h=H::new("a\\sqrt{}b");
    let radicand=h.tree()["entries"][1]["node"]["radicand"]["id"].clone();
    h.set(caret(row(&radicand,0)));
    h.edit(json!({"kind":"deleteForward"}));
    assert_eq!(h.src(),"ab");
    assert_eq!(h.sel(),caret(row(&json!(1),1)));
}

#[test]
fn backspace_at_the_start_of_a_filled_slot_still_only_moves() {
    let mut h=H::new("x^{2}");
    let sup=h.tree()["entries"][0]["node"]["slots"][0]["argument"]["id"].clone();
    h.set(caret(row(&sup,0)));
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"x^{2}");
    let mut h=H::new("\\frac{a}{}");
    let den=h.tree()["entries"][0]["node"]["denominator"]["id"].clone();
    h.set(caret(row(&den,0)));
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"\\frac{a}{}","a fraction with content is never removed by a slot backspace");
}

#[test]
fn a_range_leaving_a_matrix_edits_the_whole_highlighted_matrix() {
    let mut h=H::new(r"\begin{pmatrix}a&b\\c&d\end{pmatrix}+y");
    let grid=h.tree()["entries"][0]["node"].clone();
    assert_eq!(grid["kind"],"grid");
    let cell=grid["rows"][1]["cells"][1]["id"].clone();
    h.act_ok(json!({"kind":"pointerSelection","anchor":row(&cell,0),"focus":row(&json!(1),2)}));
    assert_eq!(h.editor()["normalizedSelection"],json!({"kind":"row","nodeId":1,"start":0,"end":2,"promoted":true}));
    h.edit(json!({"kind":"deleteBackward"}));
    assert_eq!(h.src(),"y");
    // Shift+arrow from outside into the matrix, then typing replaces it.
    let mut h=H::new(r"x+\begin{pmatrix}a&b\\c&d\end{pmatrix}");
    h.key("left",true);
    h.key("left",true);
    h.edit(json!({"kind":"insertText","text":"z","mode":"math"}));
    assert_eq!(h.src(),"x+z");
}

#[test]
fn a_raw_range_across_two_cells_of_one_matrix_is_still_refused() {
    let mut h=H::new(r"\begin{pmatrix}a&b\\c&d\end{pmatrix}");
    let grid=h.tree()["entries"][0]["node"].clone();
    let (a,d)=(grid["rows"][0]["cells"][0]["id"].clone(),grid["rows"][1]["cells"][1]["id"].clone());
    h.set(json!({"kind":"range","anchor":row(&a,0),"focus":row(&d,1)}));
    let r=h.act(json!({"kind":"edit","command":{"kind":"deleteBackward"},"groupId":null}));
    assert_eq!(r["error"]["code"],"invalidPosition");
    assert_eq!(h.src(),r"\begin{pmatrix}a&b\\c&d\end{pmatrix}");
}

#[test]
fn down_arrow_does_not_jump_into_a_subscript_elsewhere_on_the_line() {
    let mut h=H::new("a_{1}+b+c");
    let from=row(&json!(1),3);
    for down in [true,false] {
        let r=h.query(json!({"kind":"navigation","selection":caret(from.clone()),"down":down,"desiredX":null}));
        assert_eq!(r["position"],from,"down={down}");
    }
}

#[test]
fn up_and_down_move_between_base_superscript_and_subscript() {
    let mut h=H::new("x^{2}_{1}");
    let scripts=h.tree()["entries"][0]["node"].clone();
    let (base,sup,sub)=(scripts["base"]["id"].clone(),scripts["slots"][0]["argument"]["id"].clone(),scripts["slots"][1]["argument"]["id"].clone());
    let land=|h:&mut H,from:Value,down:bool|h.query(json!({"kind":"navigation","selection":caret(from),"down":down,"desiredX":null}))["position"]["nodeId"].clone();
    assert_eq!(land(&mut h,row(&base,1),false),sup);
    assert_eq!(land(&mut h,row(&base,1),true),sub);
    assert_eq!(land(&mut h,row(&sup,1),true),base);
    assert_eq!(land(&mut h,row(&sub,1),false),base);
}

#[test]
fn vertical_moves_between_lines_of_a_multiline_formula_still_work() {
    let mut h=H::new(r"\begin{gathered}a_{1}+b\\c+d\end{gathered}");
    let grid=h.tree()["entries"][0]["node"].clone();
    let (first,second)=(grid["rows"][0]["cells"][0]["id"].clone(),grid["rows"][1]["cells"][0]["id"].clone());
    let r=h.query(json!({"kind":"navigation","selection":caret(row(&first,3)),"down":true,"desiredX":null}));
    assert_eq!(r["position"]["nodeId"],second);
    let r=h.query(json!({"kind":"navigation","selection":caret(row(&second,3)),"down":false,"desiredX":null}));
    assert_eq!(r["position"]["nodeId"],first);
}

#[test]
fn clicking_beside_a_short_line_lands_on_that_line() {
    let mut h=H::new(r"\begin{gathered}x\\a+b+c+d+e+f+g\end{gathered}");
    let grid=h.tree()["entries"][0]["node"].clone();
    let first=grid["rows"][0]["cells"][0]["id"].clone();
    let layout=h.layout();let snapshot=&layout["snapshot"];
    let carets=snapshot["carets"].as_array().unwrap().clone();
    let line1=carets.iter().find(|c|c["position"]["nodeId"]==first).unwrap();
    let y=line1["baseline"].as_f64().unwrap()-0.2;
    let width=snapshot["widthEm"].as_f64().unwrap();
    let token=layout["sceneToken"].clone();let rev=h.rev;
    for (x,offset) in [(width*0.9,1),(width*0.05,0)] {
        let r=h.ok(json!({"op":"queryScene","handle":1,"expectedRevision":rev,"sceneToken":token,"query":{"kind":"hitTest","xEm":x,"yEm":y}}));
        assert_eq!(r["result"]["position"],row(&first,offset),"x={x}");
    }
}

#[test]
fn splitting_then_merging_a_line_restores_the_plain_formula() {
    for (source,at) in [("ab",1),("a+b",3),(" x ",1)] {
        let mut h=H::new(source);
        let root=h.tree()["id"].clone();
        h.set(caret(row(&root,at)));
        h.edit(json!({"kind":"splitLine"}));
        assert!(h.src().contains("gathered"),"{source}");
        h.edit(json!({"kind":"mergeLine","direction":"left"}));
        assert_eq!(h.src(),source);
        assert_eq!(h.sel(),caret(row(&root,at)),"{source}");
    }
    // A three-line group merged once keeps its environment.
    let mut h=H::new(r"\begin{gathered}a\\b\\c\end{gathered}");
    let cell=h.tree()["entries"][0]["node"]["rows"][2]["cells"][0]["id"].clone();
    h.set(caret(row(&cell,0)));
    h.edit(json!({"kind":"mergeLine","direction":"left"}));
    assert_eq!(h.src(),r"\begin{gathered}a\\bc\end{gathered}");
}

#[test]
fn text_conversion_uses_the_characters_shown_and_refuses_structures() {
    let mut h=H::new(r"\alpha+b\,c'");
    let root=h.tree()["id"].clone();
    h.set(json!({"kind":"range","anchor":row(&root,0),"focus":row(&root,6)}));
    h.edit(json!({"kind":"insertStructure","structure":"text"}));
    assert_eq!(h.src(),r"\text{α+b c′}");
    let mut h=H::new(r"\alpha+x^2");
    let root=h.tree()["id"].clone();
    h.set(json!({"kind":"range","anchor":row(&root,0),"focus":row(&root,3)}));
    let r=h.act(json!({"kind":"edit","command":{"kind":"insertStructure","structure":"text"},"groupId":null}));
    assert_eq!(r["status"],"error","{r}");
    assert_eq!(h.src(),r"\alpha+x^2");
}

#[test]
fn collapsing_a_promoted_range_lands_at_the_edge_of_the_highlight() {
    let mut h=H::new(r"a+\frac{xy}{z}+b");
    let tree=h.tree();let root=tree["id"].clone();
    let numerator=tree["entries"][2]["node"]["numerator"]["id"].clone();
    let range=json!({"kind":"range","anchor":row(&root,0),"focus":row(&numerator,1)});
    h.set(range.clone());h.key("right",false);
    assert_eq!(h.sel(),caret(row(&root,3)));
    h.set(range);h.key("left",false);
    assert_eq!(h.sel(),caret(row(&root,0)));
    // A plain range in one row still collapses to its own ends.
    h.set(json!({"kind":"range","anchor":row(&numerator,0),"focus":row(&numerator,1)}));
    h.key("right",false);
    assert_eq!(h.sel(),caret(row(&numerator,1)));
}

#[test]
fn new_spellings_survive_structured_copy_and_paste() {
    for source in [r"\displaystyle x", r"f'~y", r"a\not=b", r"\big(x\big)", r"A\xrightarrow{f}B", r"a\pmod{p}", r"\sum_{\substack{i\\j}}a"] {
        let mut h=H::new(source);
        let root=h.tree();let n=root["entries"].as_array().unwrap().len() as u32;
        h.set(json!({"kind":"range","anchor":row(&root["id"],0),"focus":row(&root["id"],n)}));
        let rev=h.rev;
        let fragment=h.ok(json!({"op":"copySelection","handle":1,"expectedRevision":rev}))["fragment"].clone();
        assert!(fragment["structure"].is_object(),"{source}: {fragment}");
        h.set(caret(row(&root["id"],n)));
        h.edit(json!({"kind":"pasteStructured","latex":fragment["latex"],"fragment":fragment["structure"]}));
        assert_eq!(h.src(),format!("{source}{source}"));
    }
}

#[test]
fn delete_forward_in_an_empty_slot_matches_backspace() {
    for structure in ["superscript","fraction"] {
        let mut h=H::new("a");
        h.edit(json!({"kind":"insertStructure","structure":structure}));
        h.edit(json!({"kind":"deleteForward"}));
        assert_eq!(h.src(),"a","{structure}");
        assert_eq!(h.sel(),caret(row(&json!(1),1)),"{structure}");
    }
}

#[test]
fn caret_moves_can_skip_resending_the_unchanged_document() {
    let source=(0..250).map(|i|format!("x_{{{i}}}+")).collect::<String>()+"y";
    let mut h=H::new(&source);
    let mut sizes=[0usize;2];let mut times=[0u128;2];
    for (mode,omit) in [false,true].into_iter().enumerate() {
        let start=std::time::Instant::now();
        for _ in 0..50 {
            h.seq+=1;
            let request=json!({"protocol":1,"command":{"op":"editorAction","handle":1,"expectedRevision":h.rev,"eventSequence":h.seq,
                "action":{"kind":"moveHorizontal","direction":"left","extend":false},"omitUnchangedDocument":omit}}).to_string();
            let reply=h.rt.dispatch_json(&request);sizes[mode]=reply.len();
            let reply:Value=serde_json::from_str(&reply).unwrap();
            assert_eq!(reply["status"],"ok");
            assert_eq!(reply["result"]["result"]["editor"].get("document").is_some(),!omit);
            assert!(reply["result"]["result"]["editor"]["selection"].is_object());
        }
        times[mode]=start.elapsed().as_micros()/50;
    }
    eprintln!("caret move reply: full {} bytes / {} us, compact {} bytes / {} us",sizes[0],times[0],sizes[1],times[1]);
    assert!(sizes[1]*20<sizes[0],"{sizes:?}");
    // A content change always carries the new document.
    h.seq+=1;
    let request=json!({"protocol":1,"command":{"op":"editorAction","handle":1,"expectedRevision":h.rev,"eventSequence":h.seq,
        "action":{"kind":"edit","command":{"kind":"insertText","text":"z","mode":"math"},"groupId":null},"omitUnchangedDocument":true}}).to_string();
    let reply:Value=serde_json::from_str(&h.rt.dispatch_json(&request)).unwrap();
    assert_eq!(reply["result"]["result"]["editor"]["document"]["revision"],h.rev+1);
}
