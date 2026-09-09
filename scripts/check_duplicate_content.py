import sys,openpyxl,collections,json
w=openpyxl.load_workbook(sys.argv[1],read_only=True,data_only=True); s=w['SEO_Products']; h=[c.value for c in next(s.iter_rows())]; ix={x:i for i,x in enumerate(h)}; fields=['title_proposed','meta_title_seo','meta_description_seo','description_proposed']; out=[]; scope={r[3] for r in w['Revision_Log'].iter_rows(min_row=2,values_only=True) if len(r)>3 and r[3]}
for f in fields:
 d=collections.defaultdict(list)
 for r in s.iter_rows(min_row=2,values_only=True):
  if r[ix['Handle']] in scope: d[str(r[ix[f]] or '').strip().casefold()].append(r[ix['Handle']])
 for v,hs in d.items():
  if v and len(hs)>1: out.append({'field':f,'handles':hs,'value':v})
print(json.dumps({'status':'PASS' if not out else 'FAIL','duplicate_count':len(out),'duplicates':out},indent=2)); sys.exit(0 if not out else 2)
