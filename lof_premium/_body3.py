# -*- coding: utf-8 -*-
# Theme card assembly + VoS Type/Source/Domain/Quality + Footer (no Sentiment, no Actions)

    parts.append(f'''
    <div class="tc"><div class="tc-h">
    <h3><span style="color:{color}">&#9679;</span> Theme {idx+1}: {esc(name)}</h3>
    <div style="display:flex;align-items:center;gap:10px"><span class="badge {bcls}">{esc(top_vt[0])}</span>
    <span style="font-weight:700">{a['pct']}%</span><span style="font-size:13px;color:var(--text2)">{a['total']} entries</span></div></div>
    <div class="tc-b">
    <div class="pb-w"><div class="pb"><div class="pb-f" style="width:{a['pct']}%;background:{color}"></div></div><span class="pb-l">{a['pct']}%</span></div>
    <p style="color:var(--text);margin-bottom:12px;line-height:1.7">{esc(desc)}</p>
    {vtb}
    <div style="display:flex;gap:8px;margin-bottom:8px;flex-wrap:wrap">
    <span class="badge" style="background:var(--err-bg);color:var(--red)">Negative {a['neg_pct']}%</span>
    <span class="badge" style="background:var(--warn-bg);color:var(--orange)">Neutral {a['neu_pct']}%</span>
    <span class="badge" style="background:var(--ok-bg);color:var(--green)">Positive {a['pos_pct']}%</span></div>
    {qhtml}
    <div class="da-sec"><div class="da-hd">&#128202; De-averaged by GMS Band</div>
    <div class="da-grid">{bcards}</div></div>
    </div></div>''')

parts.append('</div></div>')

# VoS Type Distribution
vth = ''
for vt,cnt in vos_types.most_common():
    p=round(cnt/total*100); c=vt_colors.get(vt,'#879596')
    vth += f'<div style="flex:1;min-width:180px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><span style="font-weight:600">{esc(vt)}</span><span style="font-weight:700;color:{c}">{cnt}</span></div><div class="pb-w"><div class="pb"><div class="pb-f" style="width:{p}%;background:{c}"></div></div><span class="pb-l">{p}%</span></div></div>'

# Source
srch = ''
src_c = {'beschannel_survey':'var(--blue)','amc_anecdote':'var(--purple)'}
for s,cnt in sources.most_common():
    p=round(cnt/total*100)
    srch += f'<div style="flex:1;min-width:200px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><span style="font-weight:600">{esc(source_display.get(s,s))}</span><span style="font-weight:700;color:{src_c.get(s,"#879596")}">{cnt}</span></div><div class="pb-w"><div class="pb"><div class="pb-f" style="width:{p}%;background:{src_c.get(s,"#879596")}"></div></div><span class="pb-l">{p}%</span></div></div>'

# Domain table
domh = ''
for d,cnt in top_domains:
    p=round(cnt/total*100)
    di=[r for r in rows if d in r['domain']]
    ds=Counter(r['sentiment'] for r in di)
    dn=round(ds.get('Negative',0)/len(di)*100) if di else 0
    dp=round(ds.get('Positive',0)/len(di)*100) if di else 0
    domh += f'<tr><td><strong>{esc(d)}</strong></td><td>{cnt}</td><td>{p}%</td><td><span style="color:var(--red)">{dn}%</span>/<span style="color:var(--green)">{dp}%</span></td></tr>'

qh2=quality_dist.get('high',0); qm=quality_dist.get('medium',0); ql=quality_dist.get('low',0)

parts.append(f'''
<div class="box"><div class="box-h"><h2>VoS Type Distribution</h2></div><div class="box-b"><div style="display:flex;gap:20px;flex-wrap:wrap">{vth}</div></div></div>
<div class="box"><div class="box-h"><h2>Data Source Breakdown</h2></div><div class="box-b"><div style="display:flex;gap:20px;flex-wrap:wrap">{srch}</div></div></div>
<div class="box"><div class="box-h"><h2>Top Domains</h2></div><div class="box-b"><table class="tbl"><thead><tr><th>Domain</th><th>Count</th><th>%</th><th>Neg/Pos</th></tr></thead><tbody>{domh}</tbody></table></div></div>
<div class="box"><div class="box-h"><h2>Data Quality</h2></div><div class="box-b"><div style="display:flex;gap:20px;flex-wrap:wrap">
<div style="flex:1;min-width:180px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><span style="font-weight:600">High</span><span style="font-weight:700;color:var(--green)">{qh2}</span></div><div class="pb-w"><div class="pb"><div class="pb-f" style="width:{round(qh2/total*100)}%;background:var(--green)"></div></div><span class="pb-l">{round(qh2/total*100)}%</span></div></div>
<div style="flex:1;min-width:180px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><span style="font-weight:600">Medium</span><span style="font-weight:700;color:var(--orange)">{qm}</span></div><div class="pb-w"><div class="pb"><div class="pb-f" style="width:{round(qm/total*100)}%;background:var(--orange)"></div></div><span class="pb-l">{round(qm/total*100)}%</span></div></div>
<div style="flex:1;min-width:180px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><span style="font-weight:600">Low</span><span style="font-weight:700;color:var(--red)">{ql}</span></div><div class="pb-w"><div class="pb"><div class="pb-f" style="width:{round(ql/total*100)}%;background:var(--red)"></div></div><span class="pb-l">{round(ql/total*100)}%</span></div></div>
</div></div></div>''')

# Footer
parts.append(f'''<div class="footer"><p>AGS VoS Tool &mdash; AI Analysis Summary | AI Generated (Auto Mode) | {total} entries</p>
<p style="margin-top:4px">De-averaged insights per-theme with Marketplace + Domain distribution. &copy; Amazon AGS</p></div>
</div></div></div></body></html>''')

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(''.join(parts))
print(f'HTML saved: {html_path}')
