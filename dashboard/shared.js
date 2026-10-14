const NS="http://www.w3.org/2000/svg";
function el(t,a,p){const e=document.createElementNS(NS,t);for(const k in a)e.setAttribute(k,a[k]);if(p)p.appendChild(e);return e}
function txt(p,x,y,s,o={}){const t=el("text",Object.assign({x,y,"font-size":10,fill:"#605e5c"},o),p);t.textContent=s;return t}
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const money=v=>Math.abs(v)>=1e6?"$"+(v/1e6).toFixed(2)+"M":Math.abs(v)>=1e3?"$"+(v/1e3).toFixed(Math.abs(v)>=1e5?0:1)+"K":"$"+Math.round(v);
const pct=(v,d=1)=>(v*100).toFixed(d)+"%";
const num=v=>Math.round(v).toLocaleString("en-US");