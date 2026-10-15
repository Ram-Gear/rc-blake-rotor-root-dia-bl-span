// Spreadsheet engine: evaluates ported Excel formulas (generated from the original workbook).
(function(root){
  function XErr(code){this.code=code;}
  function makeEngine(cells, fns){
    const vals={}, inputs={};
    for(const k in cells) if(cells[k].input) inputs[k]=cells[k].n;
    let cache={};
    function c(a){
      if(a in cache){ const v=cache[a]; if(v instanceof XErr) throw v; return v; }
      const d=cells[a]; let v;
      if(!d) v=0;
      else if(d.input) v=Number(inputs[a]);
      else if(d.t!==undefined) v=new XErr("#VALUE!");
      else { try{ v=fns[a](c,M,REF); if(typeof v==="number" && !isFinite(v)) v=new XErr("#NUM!"); }catch(e){ if(e instanceof XErr) v=e; else throw e; } }
      cache[a]=v; if(v instanceof XErr) throw v; return v;
    }
    const M={atan:Math.atan,acos:x=>{if(x<-1||x>1)throw new XErr("#NUM!");return Math.acos(x);},tan:Math.tan,cos:Math.cos,sin:Math.sin};
    function REF(){throw new XErr("#REF!");}
    function pow(a,b){ return Math.pow(a,b); }
    return {
      inputs,
      set(a,v){inputs[a]=v;cache={};},
      get(a){ try{return c(a);}catch(e){ if(e instanceof XErr) return e.code; throw e; } },
      recalc(){cache={};}
    };
  }
  root.makeEngine=makeEngine;
})(typeof window!=="undefined"?window:globalThis);
