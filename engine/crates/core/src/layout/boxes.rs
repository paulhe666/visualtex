use std::collections::BTreeMap;
use std::sync::Arc;
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::fonts::{FontAsset, GlyphShape, MathMetrics};
use crate::model::{AtomClass, Node, NodeKind, MAX_DEPTH, MAX_NODES};
use crate::positions::Position;
use crate::scene::{CaretStop, NodeRegion, Rect, Rule, SceneBuilder};
use crate::session::NodeId;
use super::{Cached, Key, MathStyle};

struct LocalGlyph {font_id:String,shape:Arc<GlyphShape>,x:f64,y:f64,scale:f64}
struct Child {fragment:Arc<Fragment>,x:f64,y:f64}
pub(super) struct Fragment {
    pub node_id:NodeId, role:&'static str,
    pub advance:f64,pub ascent:f64,pub descent:f64,pub italic:f64,pub ink:Option<Rect>,pub accent:Option<f64>,
    glyphs:Vec<LocalGlyph>,rules:Vec<Rect>,artworks:Vec<crate::artwork_scene::ArtworkInstance>,carets:Vec<(Position,f64,f64,f64)>,children:Vec<Child>,
    // A phantom retains child geometry for editing but contributes no paint.
    suppress_paint:bool,
    pub column_anchors:Vec<crate::scene::GridColumnAnchor>,
    color:Option<crate::paint::RgbColor>,color_change:Option<crate::paint::RgbColor>,
}
impl Fragment {
    pub fn new(node_id:NodeId,role:&'static str)->Self {Self {node_id,role,advance:0.0,ascent:0.0,descent:0.0,italic:0.0,ink:None,accent:None,glyphs:Vec::new(),rules:Vec::new(),artworks:Vec::new(),carets:Vec::new(),children:Vec::new(),suppress_paint:false,column_anchors:Vec::new(),color:None,color_change:None}}
    fn add_ink(&mut self,r:Rect) {self.ink=Some(self.ink.map_or(r,|old|old.union(r)));}
    pub fn child(&mut self,fragment:Arc<Fragment>,x:f64,y:f64) {
        self.advance=self.advance.max(x+fragment.advance);self.ascent=self.ascent.max(fragment.ascent-y);self.descent=self.descent.max(fragment.descent+y);
        if let Some(ink)=fragment.ink {self.add_ink(ink.translated(x,y));}
        self.children.push(Child {fragment,x,y});
    }
    pub(super) fn glyph(&mut self,font:&FontAsset,shape:Arc<GlyphShape>,x:f64,y:f64,scale:f64) {
        if let Some([a,b,c,d])=shape.ink {let ink=Rect::new(a*scale+x,b*scale+y,c*scale+x,d*scale+y);self.ascent=self.ascent.max(-ink.y_min);self.descent=self.descent.max(ink.y_max);self.add_ink(ink);}
        self.glyphs.push(LocalGlyph {font_id:font.identity.id.clone(),shape,x,y,scale});
    }
    pub fn artwork(&mut self,asset:&crate::custom_symbols::SymbolAsset,index:usize,shape:Arc<crate::custom_symbols::ArtworkPath>,scale:f64){
        if shape.operation==crate::custom_symbols::ArtworkOperation::Paint {let r=shape.ink;let ink=Rect::new(r.x_min*scale,r.y_min*scale,r.x_max*scale,r.y_max*scale);self.ascent=self.ascent.max(-ink.y_min);self.descent=self.descent.max(ink.y_max);self.add_ink(ink);}
        self.artworks.push(crate::artwork_scene::ArtworkInstance{node_id:self.node_id,asset_sha256:asset.sha256().into(),path_index:index,shape,x:0.,y:0.,scale,color:None});
    }
    pub fn translate_x(&mut self,delta:f64) {
        self.advance+=delta;self.ink=self.ink.map(|r|r.translated(delta,0.0));self.accent=self.accent.map(|x|x+delta);
        for glyph in &mut self.glyphs{glyph.x+=delta;}
        for artwork in &mut self.artworks{artwork.x+=delta;}
        for rule in &mut self.rules{*rule=rule.translated(delta,0.0);}
        for (_,x,_,_) in &mut self.carets{*x+=delta;}
        for child in &mut self.children{child.x+=delta;}
        for anchor in &mut self.column_anchors{anchor.x+=delta;}
    }
    pub fn rule(&mut self,r:Rect) {self.ascent=self.ascent.max(-r.y_min);self.descent=self.descent.max(r.y_max);self.add_ink(r);self.rules.push(r);}
}
pub(super) struct Builder<'a> {
    pub tex_fonts:&'a super::tex_fonts::ResolvedTexFonts<'a>,pub script_alphabet:bool,
    pub math_font:&'a FontAsset,pub text_font:&'a FontAsset,
    pub alphabet_upright_font:Option<&'a FontAsset>,pub alphabet_italic_font:Option<&'a FontAsset>,
    pub alphabet_bold_font:Option<&'a FontAsset>,pub alphabet_bold_italic_font:Option<&'a FontAsset>,pub math:&'a MathMetrics,
    pub fallback_fonts:&'a [&'a FontAsset],pub custom_text_fonts:&'a [super::ResolvedCustomTextFont<'a>],
    pub hashes:&'a BTreeMap<NodeId,[u8;32]>,pub cache:&'a mut BTreeMap<Key,Cached>,pub built:usize,pub reused:usize,
    pub variant:crate::styles::MathVariant,
    pub(super) middle_target:Option<f64>,
}
impl<'a> Builder<'a> {
    pub fn math_glyph(&self,ch:char)->EngineResult<(&'a FontAsset,u16)> {
        if let Some(id)=self.math_font.glyph_index(ch)? {return Ok((self.math_font,id));}
        for font in self.fallback_fonts {if let Some(id)=font.glyph_index(ch)? {return Ok((*font,id));}}
        // Preserve the existing explicit missing-glyph error. Never return glyph
        // zero, a different codepoint, a browser font, or a stretched substitute.
        self.math_font.glyph_id(ch).map(|id|(self.math_font,id))
    }
    pub(super) fn product_symbol_glyph(&self,ch:char,class:AtomClass)->EngineResult<Option<(&'a FontAsset,u16)>> {
        use crate::styles::MathVariant;
        // An AMS-only binding supplies explicit symbol alternates without
        // changing a non-TeX profile\'s ordinary alphabet or operator faces.
        if self.tex_fonts.contains_key("main") { return self.tex_symbol_glyph(ch,class); }
        let cp=u32::from(ch);
        let latin=ch.is_ascii_alphabetic();
        let greek_lower=matches!(cp,0x03B1..=0x03C9|0x03D1|0x03D5|0x03D6|0x03F0|0x03F1|0x03F5);
        let greek_upper=matches!(cp,0x0391..=0x03A9|0x03F4);
        let ordinary_math_letter=class==AtomClass::Ordinary&&(latin||greek_lower||greek_upper);
        let font=match self.variant {
            // Match TeX/MathLive's normal math alphabet: Latin and lowercase
            // Greek are italic; uppercase Greek, digits and ordinary symbols
            // are upright. For punctuation/operators/relations, use the real
            // selected family whenever it actually contains the glyph.
            MathVariant::Default if class==AtomClass::Ordinary&&(latin||greek_lower)=>self.alphabet_italic_font,
            MathVariant::Default=>self.alphabet_upright_font,
            MathVariant::Upright if ordinary_math_letter||ch.is_ascii_digit()=>self.alphabet_upright_font,
            MathVariant::Italic if ordinary_math_letter=>self.alphabet_italic_font,
            MathVariant::Bold if ordinary_math_letter||ch.is_ascii_digit()=>self.alphabet_bold_font,
            MathVariant::BoldItalic if ordinary_math_letter||ch.is_ascii_digit()=>self.alphabet_bold_italic_font,
            _=>None,
        };
        let Some(font)=font else{return Ok(None)};
        Ok(font.glyph_index(ch)?.map(|id|(font,id)))
    }
    pub(super) fn horizontal_product_or_math_glyph(&self,ch:char,target:Option<f64>)->EngineResult<(&'a FontAsset,crate::stretch::HorizontalGlyph)> {
        if let Some((font,gid))=self.product_symbol_glyph(ch,AtomClass::Ordinary)? {
            let shape=font.glyph(gid)?;
            if let Some(ink)=shape.ink {
                let width=ink[2]-ink[0];
                if target.is_none_or(|needed|width+1e-9>=needed) {
                    return Ok((font,crate::stretch::HorizontalGlyph {
                        width,ink_top:ink[1],ink_bottom:ink[3],
                        attachment:font.accent_attachment(gid)?-ink[0],
                        pieces:vec![crate::stretch::HorizontalPiece{x:-ink[0],shape}],
                        target_met:true,assembled:false,overlaps:Vec::new(),minimum_overlap:0.0,
                    }));
                }
            }
        }
        if let Some(target)=target {
            let mut capability_error=None;
            for font in std::iter::once(self.math_font).chain(self.fallback_fonts.iter().copied()) {
                let Some(gid)=font.glyph_index(ch)? else {continue};
                match crate::stretch::horizontal_accent(font,gid,target) {
                    Ok(glyph)=>return Ok((font,glyph)),
                    Err(error) if error.code==ErrorCode::InvalidInput=>capability_error=Some(error),
                    Err(error)=>return Err(error),
                }
            }
            if let Some(error)=capability_error{return Err(error)}
            return Err(EngineError::new(ErrorCode::MissingFont,"No mathematical font contains the requested horizontal glyph"));
        }
        let (font,gid)=self.math_glyph(ch)?;
        let shape=font.glyph(gid)?;
        let ink=shape.ink.ok_or_else(||EngineError::new(ErrorCode::InvalidInput,"Accent glyph has no outline"))?;
        Ok((font,crate::stretch::HorizontalGlyph {width:ink[2]-ink[0],ink_top:ink[1],ink_bottom:ink[3],
            attachment:font.accent_attachment(gid)?-ink[0],pieces:vec![crate::stretch::HorizontalPiece{x:-ink[0],shape}],target_met:true,
            assembled:false,overlaps:Vec::new(),minimum_overlap:0.0}))
    }
    pub(super) fn vertical_product_or_math_glyph(&self,ch:char,class:AtomClass,target:f64,scale:f64)->EngineResult<(&'a FontAsset,crate::stretch::StretchedGlyph)> {
        if let Some(variant)=self.tex_vertical_variant(ch,target/scale)? { return Ok(variant); }
        if let Some((font,gid))=self.product_symbol_glyph(ch,class)? {
            let shape=font.glyph(gid)?;
            if let Some(ink)=shape.ink {
                let height=ink[3]-ink[1];
                if height*scale+1e-9>=target {
                    return Ok((font,crate::stretch::StretchedGlyph {
                        pieces:vec![crate::stretch::StretchPiece{shape:Arc::clone(&shape),y:-ink[1]}],
                        width:shape.advance,height,italic_correction:shape.italic_correction,
                        assembled:false,overlaps:Vec::new(),minimum_overlap:0.0,
                    }));
                }
            }
        }
        let mut capability_error=None;
        for font in std::iter::once(self.math_font).chain(self.fallback_fonts.iter().copied()) {
            let Some(gid)=font.glyph_index(ch)? else {continue};
            match crate::stretch::vertical(font,gid,target/scale) {
                Ok(glyph)=>return Ok((font,glyph)),
                Err(error) if error.code==ErrorCode::InvalidInput=>capability_error=Some(error),
                Err(error)=>return Err(error),
            }
        }
        if let Some(error)=capability_error{return Err(error)}
        Err(EngineError::new(ErrorCode::MissingFont,"No mathematical font contains the requested vertical glyph"))
    }
}
impl Builder<'_> {
    /// The right half of an aligned pair has an empty ordinary atom at its
    /// left boundary. This is layout context, never a synthesized source node.
    /// Contextual cell rows are built by their grid; children keep normal caches.
    pub(super) fn math_row(&mut self,node:&Node,style:MathStyle,cramped:bool,depth:usize,
        initial:Option<AtomClass>)->EngineResult<Fragment>{
        if depth>MAX_DEPTH {return Err(EngineError::new(ErrorCode::ResourceLimit,"Layout depth budget exceeded"));}
        let NodeKind::Row{entries,..}=&node.kind else{return Err(EngineError::new(ErrorCode::InvalidInput,"Math row requires editable Row"));};
        let scale=style.scale(self.math);let mut f=Fragment::new(node.id,if depth>0&&entries.is_empty(){"emptyRow"}else{"row"});
        let mut classes:Vec<_>=entries.iter().map(|e|if matches!(e.node.kind,NodeKind::Space{..}|NodeKind::ColorChange{..}){None}else{Some(atom_class(&e.node))}).collect();
        let atoms:Vec<_>=classes.iter().enumerate().filter_map(|(i,c)|c.map(|_|i)).collect();
        for (j,&i) in atoms.iter().enumerate(){if classes[i]==Some(AtomClass::Binary){
            let left=if j==0{initial}else{classes[atoms[j-1]]};
            let right=atoms.get(j+1).and_then(|&k|classes[k]);
            if left.is_none_or(|c|matches!(c,AtomClass::Binary|AtomClass::Relation|AtomClass::Open|AtomClass::Punctuation|AtomClass::Operator))
                ||right.is_none_or(|c|matches!(c,AtomClass::Relation|AtomClass::Close|AtomClass::Punctuation)){classes[i]=Some(AtomClass::Ordinary);}
        }}
        let mut previous=initial;let top=-self.math_font.ascent*scale;let bottom=self.math_font.descent*scale;
        f.carets.push((Position::row(node.id,0),0.0,top,bottom));
        for (i,e) in entries.iter().enumerate(){
            if let Some(current)=classes[i]{if let Some(prior)=previous{f.advance+=spacing(prior,current,style)*scale;}previous=Some(current);}
            let child=self.node(&e.node,style,cramped,depth+1)?;
            let x=f.advance;let advance=x+child.advance;f.italic=child.italic;
            if entries.len()==1{f.accent=child.accent.map(|a|a+x);}
            f.child(child,x,0.0);f.advance=advance;
            f.carets.push((Position::row(node.id,i+1),f.advance,top,bottom));
        }
        if entries.is_empty(){f.advance=0.5*scale;f.ascent=(-top).max(0.0);f.descent=bottom.max(0.0);}
        Ok(f)
    }
    pub fn node(&mut self,node:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Arc<Fragment>> {
        if depth>MAX_DEPTH {return Err(EngineError::new(ErrorCode::ResourceLimit,"Layout depth budget exceeded"));}
        let key=(node.id,style,cramped,self.variant,self.middle_target.map(f64::to_bits),self.script_alphabet);let hash=self.hashes[&node.id];
        if let Some(c)=self.cache.get(&key) {if c.fingerprint==hash {self.reused+=1;return Ok(Arc::clone(&c.fragment));}}
        self.built+=1;
        let scale=style.scale(self.math);
        let mut f=match &node.kind {
            NodeKind::Row {..}=>self.math_row(node,style,cramped,depth,None)?,
            NodeKind::CustomSymbol {asset,parts,..}=>self.custom_symbol(node.id,asset,parts,style,cramped,depth)?,
            NodeKind::Grid {environment,columns,array_spec,rows,..}=>self.grid(node.id,*environment,columns,array_spec.as_ref(),rows,style,depth)?,
            NodeKind::GridRow {cells,..}=>self.grid_row(node.id,cells,style,cramped,depth)?,
            NodeKind::Symbol {value,class,command}=> {
                let mut f=Fragment::new(node.id,"symbol");
                let alternate=command.as_deref().and_then(crate::symbol_catalog::alternate);
                let single_glyph=value.chars().count()==1;
                for ch in value.chars() {
                    let preferred=if let Some(spec)=alternate {
                        if spec.value != value { return Err(EngineError::new(ErrorCode::InvalidInput,"Symbol alternate disagrees with its semantic value")); }
                        Some(self.tex_glyph(spec.tex_font_role,spec.font_character)?.ok_or_else(||
                            EngineError::new(ErrorCode::MissingFont,format!("Symbol {} requires its declared {} font glyph",spec.command,spec.tex_font_role)))?)
                    } else {self.product_symbol_glyph(ch,*class)?};
                    let (font,gid)=if let Some(preferred)=preferred {preferred} else {
                        let rendered=if self.variant==crate::styles::MathVariant::Default {math_character(ch,*class)}
                            else {crate::styles::map_character(ch,self.variant)};
                        self.math_glyph(rendered)?
                    };
                    let shape=font.glyph(gid)?;
                    let advance=shape.advance*scale;f.italic=shape.italic_correction*scale;
                    if single_glyph {f.accent=Some(font.accent_attachment(gid)?*scale);}
                    f.glyph(font,shape,f.advance,0.0,scale);f.advance+=advance;
                }
                f
            }
            NodeKind::Text {command,text,..}=> {
                let mut f=Fragment::new(node.id,if command=="\\operatorname"{"operatorName"}else{"text"});
                let font=if command=="\\operatorname"{match self.variant {crate::styles::MathVariant::Bold=>self.alphabet_bold_font.or(self.alphabet_upright_font).unwrap_or(self.math_font),crate::styles::MathVariant::BoldItalic=>self.alphabet_bold_italic_font.or(self.alphabet_upright_font).unwrap_or(self.math_font),_=>self.alphabet_upright_font.unwrap_or(self.math_font)}}else{self.text_font};let run=font.shape_text(text)?;
                f.advance=run.advance*scale;f.ascent=font.ascent*scale;f.descent=font.descent*scale;
                let mut boundaries:Vec<_>=text.grapheme_indices(true).map(|(b,_)|b).collect();boundaries.push(text.len());
                for (i,byte) in boundaries.into_iter().enumerate() {
                    let x=if byte==text.len() {run.advance} else {run.glyphs.iter().find(|g|g.cluster as usize>=byte).map_or(run.advance,|g|g.pen_x)};
                    f.carets.push((Position::text(node.id,i),x*scale,-f.ascent,f.descent));
                }
                for g in run.glyphs {f.glyph(font,g.shape,g.x*scale,g.y*scale,scale);}f
            }
            NodeKind::Fraction {command,numerator,denominator,..}=>self.fraction(node.id,command,numerator,denominator,style,cramped,depth)?,
            NodeKind::Scripts {base,slots}=>self.scripts(node.id,base,slots,style,cramped,depth)?,
            NodeKind::Radical {degree,radicand,..}=>self.radical(node.id,degree.as_deref(),radicand,style,depth)?,
            NodeKind::Fence {left,body,right,..}=>self.fence(node.id,left,body,right,style,cramped,depth)?,
            NodeKind::MiddleDelimiter {delimiter,..}=> {
                let mut f=Fragment::new(node.id,"middleDelimiter");
                if let Some(ch)=crate::delimiters::glyph(delimiter).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported middle delimiter"))? {
                    if let Some(target)=self.middle_target {
                        let (font,glyph)=self.vertical_product_or_math_glyph(ch,AtomClass::Inner,target,scale)?;
                        let axis=self.math.axis*scale;let top=-axis-glyph.height*scale/2.0;
                        f.advance=glyph.width*scale;f.italic=glyph.italic_correction*scale;
                        for piece in glyph.pieces {f.glyph(font,piece.shape,0.0,top+piece.y*scale,scale);}
                    } else {
                        let (font,gid)=match self.product_symbol_glyph(ch,AtomClass::Inner)? {Some(value)=>value,None=>self.math_glyph(ch)?};
                        let shape=font.glyph(gid)?;
                        f.advance=shape.advance*scale;f.italic=shape.italic_correction*scale;f.glyph(font,shape,0.0,0.0,scale);
                    }
                }
                f
            }
            NodeKind::Accent {command,body,..}=>self.accent(node.id,command,body,style,cramped,depth)?,
            NodeKind::Annotation {command,above,annotation,base,..}=>self.annotation(node.id,command,*above,annotation,base,style,cramped,depth)?,
            NodeKind::LargeOperator {command,parts}=>self.large_operator(node.id,command,parts,style,cramped,depth)?,
            NodeKind::GeneralOperator {body,parts,..}=>self.general_operator(node.id,body,parts,style,cramped,depth)?,
            NodeKind::MacroCall {macro_kind,arguments,..}=>self.builtin_macro(node.id,*macro_kind,arguments,style,cramped,depth)?,
            NodeKind::Style {command,variant,body,..}=> {
                let previous=self.variant;let previous_script=self.script_alphabet;self.variant=*variant;
                self.script_alphabet=matches!(command.as_str(),"\\mathscr"|"\\symscr"|"\\symbfscr");
                let child=self.node(body,style,cramped,depth+1);self.variant=previous;self.script_alphabet=previous_script;
                let child=child?;let mut f=Fragment::new(node.id,"style");
                f.italic=child.italic;f.accent=child.accent;f.child(child,0.0,0.0);f
            }
            NodeKind::Color {spec,body,..}=> {
                let child=self.node(body,style,cramped,depth+1)?;
                let mut f=Fragment::new(node.id,"color");f.color=Some(spec.color);
                f.italic=child.italic;f.accent=child.accent;f.child(child,0.0,0.0);f
            }
            NodeKind::ColorChange {spec}=> {
                let mut f=Fragment::new(node.id,"colorChange");f.color_change=Some(spec.color);f
            }
            NodeKind::Boxed {body,..}=> {
                // AMS boxed bodies are display-style even inside a script.
                // Padding is a documented engine default; rule thickness comes
                // from the active math font, not screen pixels or CSS borders.
                let child=self.node(body,MathStyle::Display,false,depth+1)?;
                let rule=self.math.fraction_rule;let gap=0.3;
                let logical=Rect::new(child.advance.min(0.0),-child.ascent,child.advance.max(0.0),child.descent);
                let bounds=child.ink.map_or(logical,|ink|logical.union(ink));
                let margin=gap+rule;let width=bounds.width()+2.0*margin;
                let top=bounds.y_min-margin;let bottom=bounds.y_max+margin;
                let mut f=Fragment::new(node.id,"boxed");
                f.child(child,margin-bounds.x_min,0.0);
                f.rule(Rect::new(0.0,top,width,top+rule));
                f.rule(Rect::new(0.0,bottom-rule,width,bottom));
                f.rule(Rect::new(0.0,top+rule,rule,bottom-rule));
                f.rule(Rect::new(width-rule,top+rule,width,bottom-rule));
                f.advance=width;f
            }
            NodeKind::Space {command,width_mu}=> {
                let mut f=Fragment::new(node.id,"space");
                f.advance=f64::from(*width_mu)/18.0*crate::spacing::space_scale(command,scale);f
            }
            NodeKind::Phantom {axis,body,..}=> {
                use crate::spacing::PhantomAxis;
                let child=self.node(body,style,cramped,depth+1)?;
                let mut f=Fragment::new(node.id,"phantom");
                f.advance=if *axis==PhantomAxis::Vertical {0.0}else{child.advance};
                if *axis!=PhantomAxis::Horizontal {f.ascent=child.ascent;f.descent=child.descent;}
                // Do not call child(): its ink and unused dimension must not
                // enlarge the phantom. Child carets/regions remain navigable.
                f.children.push(Child{fragment:child,x:0.0,y:0.0});
                f.suppress_paint=true;f
            }
            NodeKind::EmptySlot {..}=> {let mut f=Fragment::new(node.id,"emptySlot");f.advance=0.5*scale;f.ascent=self.math_font.ascent*scale;f.descent=self.math_font.descent*scale;f}
            NodeKind::UnparsedSource {..}=>return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Opaque source cannot be typeset by K0")),
        };
        f.ascent=f.ascent.max(0.0);f.descent=f.descent.max(0.0);
        if ![f.advance,f.ascent,f.descent,f.italic].into_iter().all(f64::is_finite) || f.advance.abs()>1_000_000.0 {
            return Err(EngineError::new(ErrorCode::ResourceLimit,"Layout geometry exceeds its finite extent budget"));
        }
        let f=Arc::new(f);if self.cache.len()>=MAX_NODES*4 {self.cache.clear();}
        self.cache.insert(key,Cached {fingerprint:hash,fragment:Arc::clone(&f)});Ok(f)
    }
}
fn atom_class(node:&Node)->AtomClass {match &node.kind {NodeKind::LargeOperator {..}|NodeKind::GeneralOperator {..}=>AtomClass::Operator,NodeKind::Symbol {class,..}=>*class,NodeKind::CustomSymbol {asset,..}=>asset.role(),NodeKind::Text {command,..} if command=="\\operatorname"=>AtomClass::Operator,NodeKind::MacroCall {macro_kind:crate::macros::BuiltinMacroKind::MathRelation,..}=>AtomClass::Relation,NodeKind::MacroCall {macro_kind:crate::macros::BuiltinMacroKind::MathBinary,..}=>AtomClass::Binary,NodeKind::Fraction {..}|NodeKind::Fence {..}|NodeKind::Grid {..}|NodeKind::MacroCall {..}=>AtomClass::Inner,NodeKind::Scripts {base,..}|NodeKind::Accent {body:base,..}|NodeKind::Style {body:base,..}|NodeKind::Color {body:base,..}=>atom_class(base),NodeKind::Annotation {command,base,..}=>match command.as_str(){"\\stackrel"=>AtomClass::Relation,"\\stackbin"=>AtomClass::Binary,_=>atom_class(base)}, NodeKind::Row {entries,..} if entries.len()==1=>atom_class(&entries[0].node),_=>AtomClass::Ordinary}}
fn math_character(c:char,class:AtomClass)->char {
    if class!=AtomClass::Ordinary {return c;}
    crate::styles::default_math_character(c)
}
fn spacing(a:AtomClass,b:AtomClass,style:MathStyle)->f64 {
    use AtomClass::*;
    let idx=|a|match a {Ordinary|Number=>0,Operator=>1,Binary=>2,Relation=>3,Open=>4,Close=>5,Punctuation=>6,Inner=>7};
    const MU:[[u8;8];8]=[[0,3,4,5,0,0,0,3],[3,3,0,5,0,0,0,3],[4,4,0,0,4,0,0,4],[5,5,0,0,5,0,0,5],[0,0,0,0,0,0,0,0],[0,3,4,5,0,0,0,3],[3,3,0,3,3,3,3,3],[3,3,4,5,3,0,3,3]];
    if matches!(style,MathStyle::Script|MathStyle::ScriptScript) && !matches!((a,b),(Ordinary|Number|Close|Inner,Operator)|(Operator,Ordinary|Number|Operator)) {return 0.0;}
    f64::from(MU[idx(a)][idx(b)])/18.0
}
pub(super) fn flatten(f:&Fragment,x:f64,y:f64,depth:usize,scene:&mut SceneBuilder)->EngineResult<()> {
    flatten_paint(f,x,y,depth,true,None,scene)
}
fn flatten_paint(f:&Fragment,x:f64,y:f64,depth:usize,paint:bool,color:Option<crate::paint::RgbColor>,scene:&mut SceneBuilder)->EngineResult<()> {
    let paint=paint&&!f.suppress_paint;
    let color=f.color.or(color);
    if depth>MAX_DEPTH || scene.snapshot.glyphs.len()+f.glyphs.len()>16_384 || scene.snapshot.carets.len()+f.carets.len()>32_768 {
        return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene budget exceeded"));
    }
    for anchor in &f.column_anchors {let mut anchor=anchor.clone();anchor.x+=x;scene.snapshot.grid_columns.push(anchor);}
    scene.snapshot.regions.push(NodeRegion {node_id:f.node_id,role:f.role.into(),logical:Rect::new(x.min(x+f.advance),y-f.ascent,x.max(x+f.advance),y+f.descent),ink:if paint {f.ink.map(|r|r.translated(x,y))}else{None},depth});
    if paint {
        for g in &f.glyphs {scene.glyph(f.node_id,&g.font_id,Arc::clone(&g.shape),g.x+x,g.y+y,g.scale,color);}
        for r in &f.rules {scene.snapshot.paint_order.push(crate::artwork_scene::PaintItem::Rule{index:scene.snapshot.rules.len()});scene.snapshot.rules.push(Rule {node_id:f.node_id,rect:r.translated(x,y),color});}
        for artwork in &f.artworks {
            scene.artwork_ops+=artwork.shape.path.len();
            if scene.artwork_ops>131072||scene.snapshot.artworks.len()>=4096{return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene custom artwork budget exceeded"));}
            let mut artwork=artwork.clone();artwork.x+=x;artwork.y+=y;artwork.color=color;
            scene.snapshot.paint_order.push(crate::artwork_scene::PaintItem::Artwork{index:scene.snapshot.artworks.len()});scene.snapshot.artworks.push(artwork);
        }
    }
    for (position,cx,top,bottom) in &f.carets {scene.snapshot.carets.push(CaretStop {position:*position,x:x+cx,baseline:y,top:y+top,bottom:y+bottom,depth});}
    let mut current_color=color;
    for child in &f.children {
        flatten_paint(&child.fragment,x+child.x,y+child.y,depth+1,paint,current_color,scene)?;
        // A leaf declaration updates its owning row only. Changes inside a
        // nested row/wrapper never escape to this row's following siblings.
        if let Some(next)=child.fragment.color_change {current_color=Some(next);}
    }Ok(())
}
