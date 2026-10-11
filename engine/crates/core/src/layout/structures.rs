//! Vertical constraints come from the selected font's MATH table, not CSS.
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Node,NodeKind,ScriptKind,ScriptSlot};
use crate::scene::Rect;
use crate::session::NodeId;
use super::boxes::{Builder,Fragment};
use super::MathStyle;

impl Builder<'_> {
    pub fn fraction(&mut self,id:NodeId,command:&str,numerator:&Node,denominator:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment> {
        let spec=crate::fractions::spec(command).ok_or_else(||
            EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported fraction command"))?;
        let style=spec.forced_style.unwrap_or(style);
        match spec.form {
            crate::fractions::FractionForm::Binomial=>return self.binomial(id,numerator,denominator,style,cramped,depth),
            crate::fractions::FractionForm::Bevelled=>return self.bevelled_fraction(id,numerator,denominator,style,cramped,depth),
            crate::fractions::FractionForm::Barred=>{},
        }
        let scale=style.scale(self.math);let display=style==MathStyle::Display;
        let child_style=if spec.display_arguments {MathStyle::Display} else {style.fraction_child()};
        let num=self.node(numerator,child_style,cramped,depth+1)?;
        let den=self.node(denominator,child_style,true,depth+1)?;
        let m=self.math;let axis=m.axis*scale;let thickness=m.fraction_rule*scale;
        let num_gap=(if display {m.numerator_display_gap} else {m.numerator_gap})*scale;
        let den_gap=(if display {m.denominator_display_gap} else {m.denominator_gap})*scale;
        let u=((if display {m.numerator_display_shift} else {m.numerator_shift})*scale).max(num.descent+axis+thickness/2.0+num_gap);
        let v=((if display {m.denominator_display_shift} else {m.denominator_shift})*scale).max(den.ascent-axis+thickness/2.0+den_gap);
        // Explicit horizontal policy: rule extends 0.1 local em per argument
        // edge. This does not replace font-provided vertical metrics.
        let width=num.advance.max(den.advance)+0.2*scale;
        let mut f=Fragment::new(id,"fraction");
        let nx=(width-num.advance)/2.0;let dx=(width-den.advance)/2.0;
        f.child(num,nx,-u);f.child(den,dx,v);f.advance=width;
        f.rule(Rect::new(0.0,-axis-thickness/2.0,width,-axis+thickness/2.0));Ok(f)
    }
    pub fn scripts(&mut self,id:NodeId,base:&Node,slots:&[ScriptSlot],style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment> {
        let scale=style.scale(self.math);
        let base_box=self.node(base,style,cramped,depth+1)?;
        let mut sup=None;let mut sub=None;
        for slot in slots {
            let b=self.node(&slot.argument,style.script_child(),slot.script==ScriptKind::Subscript||cramped,depth+1)?;
            if slot.script==ScriptKind::Superscript {sup=Some(b);} else {sub=Some(b);}
        }
        if brace_operator(base) { return Ok(self.stack(id,"scripts",base_box,sup,sub,style,false)); }
        let m=self.math;
        let (u,v)=self.script_shifts(&base_box,sup.as_deref(),sub.as_deref(),style,cramped,single_symbol(base));
        let bx=base_box.advance;let italic=base_box.italic;
        let mut f=Fragment::new(id,"scripts");f.accent=base_box.accent;f.child(base_box,0.0,0.0);
        if let Some(b)=sup {f.child(b,bx+italic,-u);}
        if let Some(b)=sub {f.child(b,bx,v);}
        if !slots.is_empty() {f.advance+=m.space_after_script*scale;}Ok(f)
    }
    pub(super) fn script_shifts(&self,base_box:&Fragment,sup:Option<&Fragment>,sub:Option<&Fragment>,style:MathStyle,cramped:bool,simple:bool)->(f64,f64) {
        let m=self.math;let scale=style.scale(m);
        let mut u=(if cramped {m.superscript_cramped_shift} else {m.superscript_shift})*scale;
        let mut v=m.subscript_shift*scale;
        if let Some(b)=&sup {u=u.max(b.descent+m.superscript_bottom_min*scale);if !simple {u=u.max(base_box.ascent-m.superscript_drop_max*scale);}}
        if let Some(b)=&sub {v=v.max(b.ascent-m.subscript_top_max*scale);if !simple {v=v.max(base_box.descent+m.subscript_drop_min*scale);}}
        if let (Some(a),Some(b))=(&sup,&sub) {
            let deficit=(m.sub_sup_gap*scale-(u+v-a.descent-b.ascent)).max(0.0);
            let raise=deficit.min((m.superscript_bottom_max_with_sub*scale-(u-a.descent)).max(0.0));
            u+=raise;v+=deficit-raise;
        }
        (u,v)
    }
}
fn brace_operator(node:&Node)->bool {match &node.kind {
    NodeKind::Accent {command,..}=>matches!(command.as_str(),"\\overbrace"|"\\underbrace"),
    NodeKind::Row {entries,..} if entries.len()==1=>brace_operator(&entries[0].node),_=>false,
}}
fn single_symbol(node:&Node)->bool {match &node.kind {
    NodeKind::Symbol {..}=>true,
    NodeKind::Row {entries,..} if entries.len()==1=>single_symbol(&entries[0].node),_=>false,
}}
