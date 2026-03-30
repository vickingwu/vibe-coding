# -*- coding: utf-8 -*-
# Word doc continued: value + ROI + next steps

doc.add_heading('5. 对业务的直接价值', level=1)
for t3,d in [('效率提升','从2-3天缩短到秒级自动生成'),('洞察深度','多维度De-averaged分析揭示传统方法无法发现的差异，包括Marketplace和Domain维度'),('决策精准度','基于分群洞察的差异化策略更有针对性'),('标准化输出','统一信息结构确保可比性'),('可扩展性','同一框架可应用于任意VoS数据集')]:
    p = doc.add_paragraph(); r = p.add_run(f'{t3}：'); r.bold = True; p.add_run(d)
doc.add_paragraph('')

doc.add_heading('6. 投入产出评估', level=1)
rt = doc.add_table(rows=5, cols=2, style='Light Grid Accent 1')
for i,(a,b) in enumerate([('评估项','说明'),('开发投入','约3人月'),('当前人工成本','每份报告2-3天，月均8-12人天'),('自动化节省','节省80%分析时间'),('额外价值','多维度De-averaged是人工几乎无法完成的纯增量价值')]):
    rt.rows[i].cells[0].text = a; rt.rows[i].cells[1].text = b
doc.add_paragraph('')

doc.add_heading('7. 下一步建议', level=1)
for i,s in enumerate(['将本Report作为MVP验证，收集反馈','优化De-averaged维度和展示','启动Phase 1开发','建立VoS Tool集成方案']):
    doc.add_paragraph(f'{i+1}. {s}', style='List Number')

doc.save(docx_path)
print(f'Word saved: {docx_path}')
print('All done!')
