import functools as ft
PRODUCTS=["2x4","2x6","2x8","2x10","4x4","4x6","6x6"]
LENGTHS=[8,10,12,14,16,18,20]
GRADES=["1","2","3","4","2P","MSR"]
TARGET=72; SEP="-"

def all_patterns(types):
    types=sorted(types,reverse=True); pats=[]
    def enum(i,rem,cur):
        if rem==0: pats.append(list(cur)); return
        if i>=len(types): return
        t=types[i]
        for n in range(rem//t,-1,-1): enum(i+1,rem-t*n,cur+[t]*n)
    enum(0,TARGET,[]); return pats
def dcount(p): return len(set(p))
def cmp_rows(x,y):
    m=min(len(x),len(y))
    for i in range(m):
        if x[i]!=y[i]: return -1 if x[i]>y[i] else 1
    if len(x)!=len(y): return -1 if len(x)<len(y) else 1
    return 0
def pat_better(x,y):
    if dcount(x)!=dcount(y): return dcount(x)<dcount(y)
    if len(x)!=len(y): return len(x)<len(y)
    return cmp_rows(x,y)<0

def solve_lengths(length_total,num_rows,target=72):
    counts={L:c for L,c in length_total.items() if c>0}
    types=sorted(counts,reverse=True); total=sum(L*counts[L] for L in types)
    fillable=min(num_rows,total//target); pats=all_patterns(types)
    pats.sort(key=ft.cmp_to_key(lambda x,y:-1 if pat_better(x,y) else (1 if pat_better(y,x) else 0)))
    rem=dict(counts); chosen=[]; found=[False]
    def can(p):
        nd={}
        for x in p: nd[x]=nd.get(x,0)+1
        return all(rem.get(t,0)>=n for t,n in nd.items())
    def use(p,d):
        for x in p: rem[x]=rem.get(x,0)+d
    def bt(r):
        if found[0]: return
        if r==fillable: found[0]=True; return
        for p in pats:
            if can(p):
                use(p,-1); chosen.append(list(p)); bt(r+1)
                if found[0]: return
                chosen.pop(); use(p,1)
    bt(0)
    if not found[0]: return None
    return sorted([sorted(c,reverse=True) for c in chosen], key=ft.cmp_to_key(cmp_rows))

def solve(items,num_rows,target=72):
    # items: list of dict(product,length,grade,qty)
    length_total={L:0 for L in LENGTHS}
    for it in items: length_total[it['length']]+=it['qty']
    rows=solve_lengths(length_total,num_rows,target)
    if rows is None: return None,None
    # queue per length, grouped by product order then grade order
    queue={L:[] for L in LENGTHS}
    pidx={p:i for i,p in enumerate(PRODUCTS)}; gidx={g:i for i,g in enumerate(GRADES)}
    for L in LENGTHS:
        ls=[it for it in items if it['length']==L]
        ls.sort(key=lambda it:(pidx[it['product']],gidx[it['grade']]))
        for it in ls: queue[L]+=[(it['product'],it['grade'])]*it['qty']
    qpos={L:0 for L in LENGTHS}; out=[]
    for row in rows:
        cells=[]
        for L in row:
            if qpos[L]<len(queue[L]):
                p,g=queue[L][qpos[L]]; qpos[L]+=1
                cells.append((p,L,g))
        out.append(cells)
    unplaced=[(p,L,g) for L in LENGTHS for (p,g) in queue[L][qpos[L]:]]
    return out,unplaced

DEFAULT_ITEMS=[
 dict(product="2x4",length=16,grade="2",qty=2),
 dict(product="2x6",length=16,grade="2",qty=3),
 dict(product="2x6",length=14,grade="2",qty=12),
 dict(product="2x8",length=14,grade="MSR",qty=8),
 dict(product="2x4",length=12,grade="1",qty=30),
]
if __name__=="__main__":
    rows,unp=solve(DEFAULT_ITEMS,10)
    for i,row in enumerate(rows):
        print(f"R{i+1:>2}: "+"  ".join(f"{p}{SEP}{L}{SEP}{g}" for p,L,g in row)+f"  (={sum(L for _,L,_ in row)})")
    print("Unplaced:",unp)
    from collections import Counter
    used=Counter((p,L,g) for row in rows for (p,L,g) in row)
    inv=Counter((it['product'],it['length'],it['grade']) for it in DEFAULT_ITEMS for _ in range(it['qty']))
    print("Inventory match:", used==inv, "| all 72:", all(sum(L for _,L,_ in r)==72 for r in rows))
