// One-off independent Bertha airport-source audit. Does not alter production values.
const start="2026-07-22T00:00:00Z", end="2026-07-23T23:59:00Z";
for (const station of ["MSY","NEW","HSA"]) {
 for (const types of [[3,4],[1,3,4]]) {
  const u=new URL("https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py");
  for(const [k,v] of Object.entries({station,sts:start,ets:end,tz:"Etc/UTC",format:"onlycomma",missing:"M",trace:"T",direct:"no"}))u.searchParams.set(k,v);
  for(const r of types)u.searchParams.append("report_type",String(r));
  for(const f of ["sknt","gust","peak_wind_gust","mslp"])u.searchParams.append("data",f);
  const response=await fetch(u,{headers:{"User-Agent":"PSH-test-Bertha/0.17"}});
  const raw=await response.text();
  const lines=raw.trim().split(/\r?\n/);
  const headers=(lines[0]||"").split(",");
  const idx=headers.indexOf("sknt"),gust=headers.indexOf("gust");
  const vals=lines.slice(1).map(x=>x.split(","));
  const speeds=vals.map(x=>({speed:Number(x[idx]),gust:Number(x[gust]),valid:x[headers.indexOf("valid")]}))
    .filter(x=>Number.isFinite(x.speed) && x.speed>=0);
  speeds.sort((a,b)=>b.speed-a.speed);
  console.log(JSON.stringify({station,types,status:response.status,rowCount:lines.length,
    header:headers.slice(0,14),top:speeds.slice(0,8),responseStart:raw.slice(0,180)}));
 }
}
