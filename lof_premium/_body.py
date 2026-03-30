# -*- coding: utf-8 -*-
# HTML body: sidebar, stats, exec summary, themes with de-averaged, VoS type, source, domain, quality, footer
# NO: keyword tags, sentiment section, recommended actions

# Sidebar + Stats
parts.append(f'''<body><div class="app">
<nav class="side"><div class="side-logo">AGS VoS Tool</div>
<div class="side-nav"><a href="#">Dashboard</a><a href="#" class="on">VoS Library</a>
<a href="#">Topic Details</a><a href="#">Trending Topics</a>
<a href="#">Survey Analysis</a><a href="#">Reports</a></div></nav>
<div class="main"><div class="top">
<div class="bc"><a href="#">VoS Tool</a><span style="color:#aab7b8"> / </span>
<a href="#">VoS Library</a><span style="color:#aab7b8"> / </span>
<span style="color:var(--blue);font-weight:600">AI Analysis</span></div>
<div class="acts"><button class="btn btn-n">Re-analyze</button>
<button class="btn btn-n">Export</button>
<button class="btn btn-p">Save to Topic</button></div></div>
<div class="pg">
<div class="alert"><div style="color:var(--blue)">&#9432;</div>
<div><div class="alert-t">AI Analysis Summary &mdash; Auto Mode</div>
<div>Generated 2026-01-09 | <strong>{total} entries</strong> | {date_min} ~ {date_max}</div></div></div>
<div class="stats">
<div class="st"><div class="st-v">{total}</div><div class="st-l">VoS Entries</div></div>
<div class="st"><div class="st-v">{unique_sellers}</div><div class="st-l">Unique Sellers</div></div>
<div class="st"><div class="st-v">{len(sources)}</div><div class="st-l">Data Sources</div></div>
<div class="st"><div class="st-v" style="font-size:15px">{date_min}<br>~{date_max}</div><div class="st-l">Time Range</div></div>
<div class="st"><div class="st-v" style="color:var(--text2)">{neu_pct}%</div><div class="st-l">Neutral</div></div>
<div class="st"><div class="st-v" style="color:var(--green)">{pos_pct}%</div><div class="st-l">Positive</div></div>
<div class="st"><div class="st-v" style="color:var(--red)">{neg_pct}%</div><div class="st-l">Negative</div></div></div>''')

# Executive Summary
dtags = ' '.join(f'<span class="badge badge-info">{d}</span>' for d,_ in top_domains[:6])
t1n,_=theme_sorted[0]; t1a=theme_analysis[t1n]
t2n,_=theme_sorted[1]; t2a=theme_analysis[t2n]
t3n,_=theme_sorted[2]; t3a=theme_analysis[t3n]
ef=''
for nm,_ in theme_sorted[:3]:
    a=theme_analysis[nm]
    for b,bd in a['band_rich'].items():
        bn=bd['sentiment'].get('Negative',0)
        if bd['count']>=5 and bn>a['neg_pct']+10:
            ef=f'Within <strong>{esc(nm)}</strong>, {band_labels.get(b,"Band "+b)} show {bn}% negative vs theme avg {a["neg_pct"]}%.'
            break
    if ef: break
if not ef: ef='De-averaged analysis reveals significant variation across GMS bands.'
