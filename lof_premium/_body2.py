# -*- coding: utf-8 -*-
# Executive summary + Theme cards with de-averaged (marketplace + domain, no keywords)

parts.append(f'''
<div class="box"><div class="box-h"><h2>Executive Summary</h2><span class="badge badge-info">AI Generated</span></div>
<div class="box-b">
<p style="font-size:15px;line-height:1.8">This analysis covers <strong>{total} VoS entries</strong> from <strong>{unique_sellers} unique sellers</strong> ({date_min} to {date_max}), sourced from BesChannel ({sources.get("beschannel_survey",0)}, {round(sources.get("beschannel_survey",0)/total*100)}%) and AMC ({sources.get("amc_anecdote",0)}, {round(sources.get("amc_anecdote",0)/total*100)}%).</p>
<p style="font-size:15px;line-height:1.8;margin-top:12px">Feedback clusters into <strong>six themes</strong>, led by <strong>{esc(t1n)}</strong> ({t1a["pct"]}%), <strong>{esc(t2n)}</strong> ({t2a["pct"]}%), and <strong>{esc(t3n)}</strong> ({t3a["pct"]}%). Overall: Neutral {neu_pct}%, Positive {pos_pct}%, <strong style="color:var(--red)">Negative {neg_pct}%</strong>.</p>
<p style="font-size:15px;line-height:1.8;margin-top:12px"><strong>Key de-averaged finding:</strong> {ef}</p>
<div style="margin-top:14px;display:flex;gap:6px;flex-wrap:wrap">{dtags}</div>
</div></div>''')

# Theme Clustering
parts.append(f'''<div class="box"><div class="box-h"><h2>Theme Clustering with De-averaged Insights</h2><span class="badge badge-neutral">{len(theme_sorted)} themes</span></div><div class="box-b">''')

for idx,(name,items) in enumerate(theme_sorted):
    a = theme_analysis[name]
    color = tc[idx%len(tc)]
    bcls = tbc[idx%len(tbc)]
    top_vt = a['vos_types'][0] if a['vos_types'] else ('N/A',0)
    desc = theme_descriptions.get(name,'')
    qhtml = ''.join(f'<div class="quote">{esc(q)}</div>' for q in a['quotes'])
    # VoS type badges at theme level
    vtb = '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px">'
    for vn,vc2 in a['vos_types']:
        vp = a['vos_type_pcts'].get(vn,0)
        vtb += f'<span class="badge" style="background:{vt_colors.get(vn,"#879596")}22;color:{vt_colors.get(vn,"#879596")}">{esc(vn)} {vp}%</span>'
    vtb += '</div>'

    # De-averaged cards per band
    bcards = ''
    for band in ['A','B','C','D','E','G']:
        if band not in a['band_rich']: continue
        bd = a['band_rich'][band]
        label = band_labels.get(band, f'Band {band}')
        cnt = bd['count']
        # VoS Type mini bar
        vtbar = '<div class="da-row"><span class="da-row-label">VoS Type</span><div class="da-mini-bar" style="width:120px">'
        for vn in ['Pain point','Practice Sharing','Feature Request','Positive Experience','Knowledge Gap']:
            vp = bd['vos_type'].get(vn,0)
            if vp>0: vtbar += f'<div style="width:{max(vp,2)}%;background:{vt_colors.get(vn,"#879596")}" title="{esc(vn)} {vp}%"></div>'
        vtn = bd['vos_type_top'][0]
        vtbar += f'</div><span style="font-size:12px;color:{vt_colors.get(vtn,"#879596")}">{esc(vtn)} {bd["vos_type"].get(vtn,0)}%</span></div>'
        # Sentiment mini bar
        sn=bd['sentiment'].get('Negative',0); su=bd['sentiment'].get('Neutral',0); sp=bd['sentiment'].get('Positive',0)
        sbar = f'<div class="da-row"><span class="da-row-label">Sentiment</span><div class="da-mini-bar" style="width:120px"><div style="width:{max(sn,1)}%;background:var(--red)"></div><div style="width:{max(su,1)}%;background:var(--orange)"></div><div style="width:{max(sp,1)}%;background:var(--green)"></div></div><span style="font-size:12px"><span style="color:var(--red)">{sn}%</span>/<span style="color:var(--green)">{sp}%</span></span></div>'
        # Neg flag
        nf = ''
        if sn>a['neg_pct']+8 and cnt>=5: nf=f' <span style="background:var(--err-bg);color:var(--red);padding:1px 6px;border-radius:3px;font-size:11px;font-weight:600">&#9650;+{sn-a["neg_pct"]}pp</span>'
        elif sn<a['neg_pct']-8 and cnt>=5: nf=f' <span style="background:var(--ok-bg);color:var(--green);padding:1px 6px;border-radius:3px;font-size:11px;font-weight:600">&#9660;{sn-a["neg_pct"]}pp</span>'
        # Marketplace distribution (NEW - replaces keywords)
        mph = '<div class="da-row"><span class="da-row-label">Marketplace</span><div class="da-dist">'
        for mn,mc in bd.get('mp_top',[])[:4]:
            mph += f'<span class="da-dist-item">{esc(mn)} {round(mc/cnt*100)}%</span>'
        mph += '</div></div>'
        # Domain distribution (NEW - replaces keywords)
        dmh = '<div class="da-row"><span class="da-row-label">Domain</span><div class="da-dist">'
        for dn,dc in bd.get('dom_top',[])[:4]:
            dmh += f'<span class="da-dist-item">{esc(dn)} {round(dc/cnt*100)}%</span>'
        dmh += '</div></div>'
        # Concentration + Quality
        mh = f'<div class="da-row"><span class="da-row-label">Concentration</span><span style="font-size:12px;font-weight:700">{bd["concentration"]}% of theme</span></div>'
        mh += f'<div class="da-row"><span class="da-row-label">High Quality</span><span style="font-size:12px">{bd["high_quality_pct"]}%</span></div>'
        # Quote
        qh = f'<div class="da-quote">{esc(bd["quote"])}</div>' if bd['quote'] else ''
        bcards += f'''<div class="da-card"><div class="da-card-top"><span class="da-card-name">{esc(label)}{nf}</span><span class="da-card-meta">{cnt} entries</span></div>{vtbar}{sbar}{mph}{dmh}{mh}{qh}</div>'''
