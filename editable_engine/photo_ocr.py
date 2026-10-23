"""Embedded PP-OCR for photographed grids; no network or model downloads."""
from __future__ import annotations
import re
import cv2
import numpy as np

_engine=None


def recognizer():
    global _engine
    if _engine is None:
        import onnxruntime
        onnxruntime.disable_telemetry_events()
        from rapidocr_onnxruntime import RapidOCR
        _engine=RapidOCR(intra_op_num_threads=2,inter_op_num_threads=1,det_limit_side_len=1800)
    return _engine


def clean_identifier(text):
    text=re.sub(r'\s+','',text).strip('.,，。_ |：:')
    text=re.sub(r'[~～—–−]','-',text)
    text=re.sub(r'-{2,}','-',text)
    match=re.search(r'[A-Z]{4}\d{4}-\d{2}-\d{2}-\d{3}(?!\d)',text)
    if match:return match.group()
    text=re.sub(r'(M\d+)[Xx](\d)',r'\1×\2',text)
    if re.fullmatch(r'[A-Z]+\d+\.\d+-M\d+×\d+\.\d+',text):return text
    return None


def recognize_grid(image, table, fallback, repair):
    """Detect tight printed text boxes, then assign them by measured cells."""
    engine=recognizer()
    results,_=engine(image,use_cls=False)
    xs,ys=table['xs'],table['ys']
    rows,cols=len(ys)-1,len(xs)-1
    cells=[[[] for _ in range(cols)] for _ in range(rows)]
    for box,text,score in results or []:
        points=np.asarray(box)
        cx,cy=np.mean(points,axis=0)
        row,col=int(np.searchsorted(ys,cy)-1),int(np.searchsorted(xs,cx)-1)
        if 0<=row<rows and 0<=col<cols and score>=.65:
            cells[row][col].append((str(text),float(score),points))
    # Empty cells may still have a reliable WinRT reading; never fill a value
    # from row order, neighbouring rows or a calculated total.
    for line in fallback.get('lines',[]):
        for word in line.get('words',[]):
            cx=float(word['x'])+float(word.get('w',0))/2
            cy=float(word['y'])+float(word.get('h',0))/2
            row,col=int(np.searchsorted(ys,cy)-1),int(np.searchsorted(xs,cx)-1)
            if 0<=row<rows and 0<=col<cols and not cells[row][col]:
                if re.fullmatch(r'\d+(?:[.,]\d+)?',str(word.get('t',''))):
                    cells[row][col].append((str(word['t']),.7,None))
    unresolved=[]
    words=[]
    for row in range(rows):
        for col in range(cols):
            candidates=cells[row][col]
            if not candidates:
                unresolved.append([row,col])
                continue
            text,score,box=max(candidates,key=lambda entry: entry[1]*max(1,len(entry[0]))**.25)
            value=repair(text).strip(' |')
            header=''.join(entry[0] for entry in cells[0][col])
            if row>0 and '图号' in header:
                normalized=clean_identifier(value)
                if normalized is None and box is not None:
                    x0,y0=np.floor(box.min(axis=0)).astype(int)
                    x1,y1=np.ceil(box.max(axis=0)).astype(int)
                    crop=image[max(0,y0):y1,max(0,x0):x1]
                    variants=[]
                    for trim in (0,.025,.05,.075):
                        end=max(10,round(crop.shape[1]*(1-trim)))
                        for inset in (0,2,4):
                            if crop.shape[0]>inset*2+8:
                                variants.append(crop[inset:crop.shape[0]-inset,:end])
                    readings,_=engine.text_rec(variants)
                    valid=[(s,clean_identifier(t)) for t,s in readings if clean_identifier(t)]
                    if valid:normalized=max(valid,key=lambda entry: entry[0])[1]
                if normalized:value=normalized
                else:unresolved.append([row,col])
            if row>0 and any(label in header for label in ('序号','数量','总价','金额')):
                value=value.strip('.,，。_ ') 
                if not re.fullmatch(r'\d+(?:\.\d+)?',value):unresolved.append([row,col])
            words.append({'t':value,'x':float(xs[col]+4),'y':float(ys[row]+8),
                          'w':float(xs[col+1]-xs[col]-8),'h':float(ys[row+1]-ys[row]-16)})
    return {'ok':True,'width':image.shape[1],'height':image.shape[0],
            'lines':[{'text':word['t'],'words':[word]} for word in words],
            'photo_unresolved_cells':unresolved,'photo_ocr_engine':'embedded-ppocr'}
