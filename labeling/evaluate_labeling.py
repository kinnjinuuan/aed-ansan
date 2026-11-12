"""라벨링 평가: 팀 정답 vs AI 분류.
사용: python evaluate_labeling.py 라벨링시트_200건.xlsx 표본키_비공개.csv
시트에 '정답라벨' 열(팀 합의)이 있으면 그것을, 없으면 라벨_담당자A를 정답으로 사용.
"""
import sys, pandas as pd
sheet, keyf = sys.argv[1], sys.argv[2]
d = pd.read_excel(sheet, sheet_name='라벨링'); k = pd.read_csv(keyf)
d = d.merge(k, on='표본번호')
truth = '정답라벨' if '정답라벨' in d else '라벨_담당자A'
d = d.dropna(subset=[truth])
for name in ['AI라벨', '등록시간기준라벨']:
    sub = d[d[truth].isin(['24시간 가능','업무시간 한정','판단불가'])]
    ok = (sub[name] == sub[truth])
    print(f'\n[{name}] 정확도 {ok.mean():.1%} (n={len(sub)})')
    print(pd.crosstab(sub[truth], sub[name], rownames=['정답'], colnames=['예측']))
    print(sub.assign(ok=ok).groupby('시설유형')['ok'].agg(['mean','count']).round(2))
if {'라벨_담당자A','라벨_담당자B'} <= set(d.columns):
    from sklearn.metrics import cohen_kappa_score
    b = d.dropna(subset=['라벨_담당자A','라벨_담당자B'])
    print('\nCohen κ (A vs B):', round(cohen_kappa_score(b['라벨_담당자A'], b['라벨_담당자B']), 3))
