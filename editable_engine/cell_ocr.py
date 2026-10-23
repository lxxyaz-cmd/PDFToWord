"""Conservative cell-level OCR verification for ruled scans, entirely offline.

Whole-page OCR can omit short decimals or read 8 as &. Read the actual cell
pixels twice before replacing text; arithmetic is only a warning, never a
source of replacement values. Empty optional cells are not filled by a model.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
import re

import cv2
import numpy as np

from scan_tables import cell_groups, word_box, words_text


def number(text):
    """Strict, unambiguous decimal syntax (do not guess comma separators)."""
    value = str(text).strip()
    if not re.fullmatch(r'[+-]?\d+(?:\.\d+)?', value):
        return None
    try:
        return Decimal(value)
    except InvalidOperation:
        return None


def comparable(text):
    text = re.sub(r'\s+','',str(text))
    return re.sub(r'(?<=\d)[xX*×](?=\d)','×',text)


def cell_crops(image, bounds):
    x0, y0, x1, y1 = bounds
    crop = image[max(0,y0+7):y1-7, max(0,x0+7):x1-7]
    if not crop.size:
        return None
    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    yy, xx = np.where(gray < 150)
    if len(xx) < 12 or np.ptp(yy) < 4:
        return None
    tight = crop[max(0,int(yy.min())-3):int(yy.max())+4,
                 max(0,int(xx.min())-3):int(xx.max())+4]
    # A second view changes sampling and contrast, not the printed content.
    second = cv2.resize(tight, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    second = cv2.copyMakeBorder(second, 4, 4, 4, 4, cv2.BORDER_CONSTANT, value=(255,255,255))
    return tight, second


def accepted_reading(old, readings, numeric):
    """Require two high-confidence matching readings, with field constraints."""
    if len(readings) != 2 or min(float(score) for _,score in readings) < .98:
        return None
    values = [str(text).strip() for text,_ in readings]
    if values[0] != values[1]:
        return None
    value = values[0]
    if numeric:
        return value if number(value) is not None else None
    # Only repair mixed text/specifications here. Do not rewrite Chinese names
    # on the strength of a second engine, or hallucinate optional drawing IDs.
    if not old or not re.search(r'\d', value):
        return None
    if re.findall(r'[\u4e00-\u9fff]', old) != re.findall(r'[\u4e00-\u9fff]', value):
        return None
    return value


def verify_stamp_prose(image,tables,stamps,payload,repair,engine):
    """Re-read existing prose crossing a seal outside table bounds."""
    for line in payload['lines']:
        words=line.get('words',[])
        if not words:
            continue
        boxes=[word_box(w) for w in words]
        x0,y0=min(b[0] for b in boxes),min(b[1] for b in boxes)
        x1,y1=max(b[2] for b in boxes),max(b[3] for b in boxes)
        if any(t['ys'][0] <= (y0+y1)/2 <= t['ys'][-1] for t in tables):
            continue
        if not any(x0 < s['bbox'][2] and x1 > s['bbox'][0] and y0 < s['bbox'][3]
                   and y1 > s['bbox'][1] for s in stamps):
            continue
        old=words_text(words,repair)
        if len(re.findall(r'[\u4e00-\u9fff]',old)) < 4:
            continue
        views=cell_crops(image,(int(x0)-15,int(y0)-15,int(x1)+15,int(y1)+15))
        if views is None:
            continue
        readings,_=engine.text_rec(list(views))
        value=str(readings[0][0]).strip() if len(readings)==2 else ''
        accepted=(len(readings)==2 and value==str(readings[1][0]).strip()
                  and min(float(s) for _,s in readings)>=.98
                  and SequenceMatcher(None,comparable(old),comparable(value)).ratio()>=.75
                  and re.findall(r'\d+',old)==re.findall(r'\d+',value))
        location={'table':0,'row':0,'column':0,'region':'印章附近页脚文字'}
        if accepted and comparable(value)!=comparable(old):
            payload['cell_ocr_changes'].append(dict(location,before=old,after=value))
            line['text']=value
            line['words']=[{'t':value,'x':x0,'y':y0,'w':x1-x0,'h':y1-y0}]
        elif not accepted and any(comparable(t)!=comparable(old) for t,s in readings):
            payload['cell_ocr_issues'].append(dict(location,text=old,reason='OCR readings disagree'))


def verify_cells(image, tables, payload, repair, engine=None, stamps=()):
    from photo_ocr import recognizer
    engine = engine or recognizer()
    words = [w for line in payload.get('lines',[]) for w in line.get('words',[])]
    replacements = {}
    jobs = []
    crops = []
    issues = []
    changes = []
    for ti, table in enumerate(tables):
        xs, ys = table['xs'], table['ys']
        assigned = {}
        for r0,c0,r1,c1 in cell_groups(table):
            ww = [w for w in words if xs[c0] <= (word_box(w)[0]+word_box(w)[2])/2 < xs[c1+1]
                  and ys[r0] <= (word_box(w)[1]+word_box(w)[3])/2 < ys[r1+1]]
            assigned[r0,c0] = (r1,c1,ww,words_text(ww,repair))
        labels = {c: re.sub(r'\s+','',cell[3]) for (r,c),cell in assigned.items() if r == 0}
        numeric_cols = {c for c,label in labels.items()
                        if any(s in label for s in ('序号','数量','单价','金额','合计','总价'))}
        if len(numeric_cols) < 2:
            continue
        for (r,c),(r1,c1,ww,old) in assigned.items():
            if r == 0 or r1 != r or c1 != c:
                continue
            numeric = c in numeric_cols
            # Summary labels in a numeric column are labels, not missing values.
            if numeric and any(label in old for label in ('合计','金额','总计','小计')):
                continue
            if not numeric and (not old or not re.search(r'\d',old)):
                continue
            views = cell_crops(image,(xs[c],ys[r],xs[c+1],ys[r+1]))
            if views is None:
                continue
            jobs.append((ti,r,c,ww,old,numeric,xs[c],ys[r],xs[c+1],ys[r+1]))
            crops.extend(views)
    # Bounded batches avoid retaining model intermediate arrays for a long PDF.
    readings = []
    for start in range(0,len(crops),64):
        batch,_ = engine.text_rec(crops[start:start+64])
        readings.extend(batch)
    if len(readings) != len(crops):
        raise ValueError('Cell OCR returned an incomplete batch')
    for i,(ti,r,c,ww,old,numeric,x0,y0,x1,y1) in enumerate(jobs):
        pair = readings[2*i:2*i+2]
        value = accepted_reading(old,pair,numeric)
        if value is not None and (value != old if numeric else comparable(value) != comparable(old)):
            # Keep existing layout and all words outside this measured cell.
            for word in ww:
                replacements[id(word)] = None
            changes.append({'table':ti+1,'row':r+1,'column':c+1,'before':old,'after':value})
            replacements[('new',i)] = {'t':value,'x':float(x0+8),'y':float(y0+8),
                                      'w':float(x1-x0-16),'h':float(y1-y0-16)}
        if value is None and (numeric and (number(old) is None or any(number(v) != number(old) for v,s in pair))
                              or not numeric and any(comparable(v) != comparable(old) for v,s in pair)):
            issues.append({'table':ti+1,'row':r+1,'column':c+1,'text':old,'reason':'OCR readings disagree'})
    retained = []
    for line in payload.get('lines',[]):
        kept = [w for w in line.get('words',[]) if id(w) not in replacements]
        if kept:
            retained.append(dict(line,words=kept,text=words_text(kept,repair)))
    retained.extend({'text':w['t'],'words':[w]} for key,w in replacements.items() if isinstance(key,tuple))
    payload['lines'] = retained
    payload['cell_ocr_changes'] = changes
    payload['cell_ocr_issues'] = issues
    payload['cell_ocr_checked'] = len(jobs)
    # Cross-check numeric relationships after recognition. A mismatch is only
    # a review flag: source quotations can contain discounts or original errors.
    updated = [w for line in retained for w in line['words']]
    for ti,table in enumerate(tables):
        xs,ys = table['xs'],table['ys']
        def text_at(r,c):
            return words_text([w for w in updated if xs[c] <= w['x']+w['w']/2 < xs[c+1]
                               and ys[r] <= w['y']+w['h']/2 < ys[r+1]],repair)
        headers = [text_at(0,c) for c in range(len(xs)-1)]
        quantity = next((c for c,v in enumerate(headers) if '数量' in v),None)
        price = next((c for c,v in enumerate(headers) if '单价' in v),None)
        amount = next((c for c,v in enumerate(headers) if any(s in v for s in ('合计','金额','总价'))),None)
        if None in (quantity,price,amount):
            continue
        for r in range(1,len(ys)-1):
            q,p,a = (number(text_at(r,c)) for c in (quantity,price,amount))
            if None not in (q,p,a) and abs(q*p-a) > Decimal('.01'):
                issues.append({'table':ti+1,'row':r+1,'column':amount+1,
                               'text':text_at(r,amount),'reason':'quantity times price differs from amount'})
    verify_stamp_prose(image,tables,stamps,payload,repair,engine)
    return payload
