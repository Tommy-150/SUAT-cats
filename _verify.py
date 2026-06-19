import sys, json, tempfile, os
sys.path.insert(0, 'D:/Documents/OpenHanako/SUAT-cats')
from pathlib import Path
import cats_manager as m

store = m.ExcelStore(m.EXCEL_PATH)
store.load()
rows = store.all_rows()

tmp = Path(tempfile.gettempdir()) / "cats_test.json"
count, warns = m.regenerate_cats_json(rows, m.CLASSIFIED_DIR, tmp)
print('generated:', count, 'cats')
print('warnings:', len(warns))
for w in warns[:5]:
    print('  ', w)

old = json.loads(open(m.JSON_PATH, encoding='utf-8').read())
new = json.loads(open(tmp, encoding='utf-8').read())
print('len old/new:', len(old), len(new))

# \u9010\u53ea\u5bf9\u6bd4\u4e3b\u8981\u5b57\u6bb5
diff_count = 0
for o, n in zip(old, new):
    for k in ('id','name','gender','affection','status','desc','story','avatar','avatar_hd'):
        if o.get(k) != n.get(k):
            print('DIFF', o['id'], k, '|', repr(o.get(k))[:60], '!=', repr(n.get(k))[:60])
            diff_count += 1
    if len(o.get('otherPhotos',[])) != len(n.get('otherPhotos',[])):
        print('DIFF', o['id'], 'otherPhotos count', len(o.get('otherPhotos',[])), '!=', len(n.get('otherPhotos',[])))
        diff_count += 1
print('total diffs:', diff_count)
os.unlink(tmp)
