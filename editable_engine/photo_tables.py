"""Conservative grid recovery for a photographed, folded ruled document."""
from __future__ import annotations
import cv2
import numpy as np


def _peaks(values, threshold):
    groups=[]
    for value in np.where(values >= threshold)[0]:
        if not groups or value-groups[-1][-1]>4:
            groups.append([int(value)])
        else:
            groups[-1].append(int(value))
    return [round(float(np.mean(group))) for group in groups]


def _regular_run(peaks):
    best=[]
    for start in range(len(peaks)):
        run=peaks[start:start+1]
        for point in peaks[start+1:]:
            gap=point-run[-1]
            pitch=float(np.median(np.diff(run))) if len(run)>1 else gap
            if gap < 15 or gap > 70 or (len(run)>2 and not .68*pitch <= gap <= 1.4*pitch):
                break
            run.append(point)
        if len(run)>len(best):best=run
    return best


def _trace_rows(ink, seeds, reference, left, right):
    """Trace ink seams from a known regular column, permitting folded curves."""
    h,w=ink.shape
    score=cv2.boxFilter((ink>0).astype(np.float32),-1,(11,1),normalize=True)
    result=np.zeros((len(seeds),right-left+1),np.float32)
    for index,seed in enumerate(seeds):
        result[index,reference-left]=seed
        for direction in (-1,1):
            y=float(seed)
            velocity=0.
            for x in range(reference+direction,left-1 if direction<0 else right+1,direction):
                predicted=y+velocity
                candidates=np.arange(max(2,round(y)-3),min(h-2,round(y)+4))
                evidence=np.maximum.reduce([score[candidates-1,x],score[candidates,x],score[candidates+1,x]])
                cost=evidence - .055*np.abs(candidates-predicted) - .008*np.abs(candidates-seed)
                selected=float(candidates[int(np.argmax(cost))])
                velocity=.8*velocity+.2*(selected-y)
                y=selected
                result[index,x-left]=y
    return result


def rectify_photo_table(image, diagnostics=None):
    """Return a flattened grid only when many regular rows and columns agree."""
    factor=min(1.,1200/image.shape[1])
    small=cv2.resize(image,None,fx=factor,fy=factor,interpolation=cv2.INTER_AREA) if factor<1 else image.copy()
    height,width=small.shape[:2]
    gray=cv2.cvtColor(small,cv2.COLOR_RGB2GRAY)
    ink=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,31,12)
    horizontal=cv2.morphologyEx(ink,cv2.MORPH_OPEN,np.ones((1,max(23,width//43)),np.uint8))
    best=[]
    reference=0
    for x in range(round(width*.12),round(width*.85),max(10,width//65)):
        peaks=_peaks(np.count_nonzero(horizontal[:,x-14:x+15],axis=1),14)
        run=_regular_run([p for p in peaks if height*.1<p<height*.94])
        if len(run)>len(best):best,reference=run,x
    if len(best)<10:
        return None
    top,bottom=best[0],best[-1]
    vertical=cv2.morphologyEx(ink,cv2.MORPH_OPEN,np.ones((max(30,height//55),1),np.uint8))
    lines=cv2.HoughLinesP(vertical,1,np.pi/180,60,minLineLength=max(160,(bottom-top)//7),maxLineGap=35)
    if lines is None:return None
    middle=(top+bottom)/2
    candidates=[]
    for x0,y0,x1,y1 in lines[:,0]:
        if abs(y1-y0)<100 or abs((x1-x0)/(y1-y0))>.15:continue
        slope=(x1-x0)/(y1-y0)
        intercept=x0-slope*y0
        center=slope*middle+intercept
        if 0<=center<width:
            candidates.append((center,float(x0),float(y0),float(x1),float(y1)))
    groups=[]
    for entry in sorted(candidates):
        if not groups or entry[0]-groups[-1][-1][0] > width*.035:groups.append([entry])
        else:groups[-1].append(entry)
    columns=[]
    for group in groups:
        xx=np.array([v for entry in group for v in (entry[1],entry[3])])
        yy=np.array([v for entry in group for v in (entry[2],entry[4])])
        fit=np.polyfit(yy,xx,1)
        x=round(float(np.polyval(fit,middle)))
        band=horizontal[top:bottom,max(0,x-12):min(width,x+13)]
        density=np.count_nonzero(band)/max(1,band.size)
        if diagnostics is not None:diagnostics.setdefault('column_candidates',[]).append((x,round(density,4),float(np.ptp(yy))))
        if density>.018 and np.ptp(yy)>(bottom-top)*.25:columns.append(fit)
    if not 3<=len(columns)<=15:return None
    left=max(0,round(min(float(np.polyval(fit,y)) for fit in columns for y in (top,bottom)))-8)
    right=min(width-1,round(max(float(np.polyval(fit,y)) for fit in columns for y in (top,bottom)))+8)
    if right-left<width*.5 or not left<reference<right:return None
    curves=_trace_rows(ink,best,reference,left,right)
    # Traces may not cross or collapse: this is a geometry check, not inferred
    # document data. Reject unreliable recovery instead of inventing rows.
    gaps=np.diff(curves,axis=0)
    if np.mean(gaps<9)>.015:return None
    boundaries=np.array([np.polyval(fit,np.array(best)) for fit in columns],np.float32).T
    # The outer border is sometimes cut by the camera frame near the bottom.
    # Keep a small observed margin so its fit cannot clip the first digit of
    # an item number. This changes the crop, never the recognised value.
    boundaries[:,0]=np.maximum(0,boundaries[:,0]-width*.02)
    source_x=np.arange(left,right+1,dtype=np.float32)
    intersections=np.array([[np.interp(boundaries[r,c],source_x,curves[r])
                             for c in range(len(columns))] for r in range(len(best))],np.float32)
    widths=np.median(np.diff(boundaries,axis=1),axis=0)
    widths=np.maximum(35,widths)
    target_x=np.r_[0,np.cumsum(widths)].astype(np.float32)
    pitch=float(np.median(np.diff(best)))
    target_y=np.arange(len(best),dtype=np.float32)*pitch
    out_w,out_h=round(float(target_x[-1])),round(float(target_y[-1]))
    map_x=np.zeros((out_h+1,out_w+1),np.float32)
    map_y=np.zeros_like(map_x)
    for r in range(len(best)-1):
        yy=np.arange(round(target_y[r]),round(target_y[r+1])+1)
        fy=(yy-target_y[r])/pitch
        for c in range(len(columns)-1):
            xx=np.arange(round(target_x[c]),round(target_x[c+1])+1)
            fx=(xx-target_x[c])/(target_x[c+1]-target_x[c])
            xt=(1-fx)*boundaries[r,c]+fx*boundaries[r,c+1]
            xb=(1-fx)*boundaries[r+1,c]+fx*boundaries[r+1,c+1]
            yt=np.interp(xt,source_x,curves[r])
            yb=np.interp(xb,source_x,curves[r+1])
            map_x[np.ix_(yy,xx)]=(1-fy[:,None])*xt+fy[:,None]*xb
            map_y[np.ix_(yy,xx)]=(1-fy[:,None])*yt+fy[:,None]*yb
    # Local illumination normalisation removes the fold shadows, keeping the
    # observed printed strokes (no text synthesis or content correction).
    background=cv2.GaussianBlur(gray,(0,0),17)
    normalized=cv2.divide(gray,background,scale=255)
    ruled=np.zeros_like(normalized)
    for curve in curves:
        points=np.c_[source_x,curve].round().astype(np.int32)
        cv2.polylines(ruled,[points],False,255,5)
    for fit in columns[1:]:
        yy=np.arange(max(0,top-50),min(height,bottom+60))
        points=np.c_[np.polyval(fit,yy),yy].round().astype(np.int32)
        cv2.polylines(ruled,[points],False,255,4)
    normalized[ruled>0]=255
    # Preserve ink artwork beside the table (e.g. an edge-clipped grey seal).
    # It stays a separate image and is never fed into text recognition.
    side_mask=np.zeros_like(ink)
    for y in range(max(0,top),min(height,bottom)):
        x=round(float(np.polyval(columns[-1],y)))+3
        side_mask[y,max(0,x):round(width*.975)]=ink[y,max(0,x):round(width*.975)]
    contours,_=cv2.findContours(cv2.dilate(side_mask,np.ones((15,15),np.uint8)),cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    artwork=[]
    for contour in contours:
        x,y,w,h=cv2.boundingRect(contour)
        if not (width*.025<w<width*.2 and 50<h<(bottom-top)*.35):continue
        count=np.count_nonzero(side_mask[y:y+h,x:x+w])
        if count<150:continue
        artwork.append((count,x,y,w,h))
    artwork=sorted(artwork,reverse=True)[:1]
    flat=cv2.remap(normalized,map_x,map_y,cv2.INTER_CUBIC,borderMode=cv2.BORDER_REPLICATE)
    flat=cv2.cvtColor(flat,cv2.COLOR_GRAY2RGB)
    # Lines were measured above. Redraw a clean grid so the ordinary Word table
    # builder can recover the exact row/column count after geometric correction.
    for x in target_x:cv2.line(flat,(round(x),0),(round(x),out_h),(0,0,0),2)
    for y in target_y:cv2.line(flat,(0,round(y)),(out_w,round(y)),(0,0,0),2)
    # The printed title follows the same folded baseline as the top border.
    # Keep it as observed artwork rather than silently dropping the heading.
    title_x=np.linspace(boundaries[0,0],boundaries[0,-1],out_w+1,dtype=np.float32)
    title_base=np.interp(title_x,source_x,curves[0]).astype(np.float32)
    title_map_x=np.broadcast_to(title_x,(55,out_w+1)).copy()
    title_map_y=title_base[None,:]+np.arange(-60,-5,dtype=np.float32)[:,None]
    title=cv2.remap(normalized,title_map_x,title_map_y,cv2.INTER_CUBIC,borderMode=cv2.BORDER_REPLICATE)
    title=cv2.cvtColor(title,cv2.COLOR_GRAY2RGB)
    flat=np.concatenate([title,np.full((10,out_w+1,3),255,np.uint8),flat],axis=0)
    flat=cv2.copyMakeBorder(flat,15,15,15,15,cv2.BORDER_CONSTANT,value=(255,255,255))
    if artwork:
        flat=cv2.copyMakeBorder(flat,0,0,0,max(entry[3] for entry in artwork)+20,cv2.BORDER_CONSTANT,value=(255,255,255))
    flat=cv2.resize(flat,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
    if diagnostics is not None:
        diagnostics.update({'rows':len(best)-1,'columns':len(columns)-1,'reference_x':reference,
                            'source_rows':best,'column_x':boundaries.tolist(),'row_y':intersections.tolist()})
        diagnostics['side_artwork']=[]
        for _,x,y,w,h in artwork:
            rgba=cv2.cvtColor(small[y:y+h,x:x+w],cv2.COLOR_RGB2BGRA)
            rgba[:,:,3]=side_mask[y:y+h,x:x+w]
            middle_row=float(np.interp(y+h/2,intersections[:,-1],np.arange(len(best))))
            target_top=round(2*(15+65+middle_row*pitch-h/2))
            target_left=round(2*(15+out_w-6))
            diagnostics['side_artwork'].append({'bbox':(target_left,target_top,target_left+2*w,target_top+2*h),
                                                'png':cv2.imencode('.png',rgba)[1].tobytes()})
    return flat
