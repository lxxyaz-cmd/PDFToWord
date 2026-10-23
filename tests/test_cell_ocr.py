"""Acceptance gates: never guess decimals, overwrite names, or fill blank IDs."""
from pathlib import Path
import sys
import unittest
import cv2
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'editable_engine'))
from cell_ocr import accepted_reading, number, comparable, verify_cells
from scan_tables import detect_tables


class AcceptanceTests(unittest.TestCase):
    def test_invalid_amount_and_missing_price_recovered_from_two_readings(self):
        for old,new in [('生50','4.50'),('& 80','8.80'),('& 00','8.00'),('','0.3'),('0，065','0.065')]:
            self.assertEqual(accepted_reading(old,[(new,.999),(new,.999)],True),new)

    def test_conflict_or_uncertainty_is_not_a_correction(self):
        self.assertIsNone(accepted_reading('8.00',[('8.00',.999),('3.00',.999)],True))
        self.assertIsNone(accepted_reading('8.00',[('3.00',.8),('3.00',.8)],True))
        self.assertIsNone(accepted_reading('8.00',[('& 00',.999),('& 00',.999)],True))
        self.assertIsNone(number('1,000'))

    def test_specs_preserve_names_and_blank_optional_cells(self):
        self.assertEqual(accepted_reading('圆螺母M25*1',[('圆螺母M25*1.5',.999)]*2,False),'圆螺母M25*1.5')
        self.assertIsNone(accepted_reading('圆螺母M25*1',[('圆螺栓M25*1.5',.999)]*2,False))
        self.assertIsNone(accepted_reading('',[('12345',.999)]*2,False))
        self.assertEqual(comparable('销轴B30 × 80'),comparable('销轴B30X80'))

    def test_arithmetic_is_warning_only_and_optional_blanks_stay_empty(self):
        image=np.full((300,800,3),255,np.uint8)
        xs,ys=[30,150,350,550,750],[30,100,180,260]
        for x in xs:cv2.line(image,(x,30),(x,260),(0,0,0),2)
        for y in ys:cv2.line(image,(30,y),(750,y),(0,0,0),2)
        words=[]
        for r,values in enumerate([['序号','数量','单价','金额'],['1','2','3','5']]):
            for c,value in enumerate(values):
                words.append({'t':value,'x':xs[c]+20,'y':ys[r]+20,'w':40,'h':30})
                if r:cv2.putText(image,value,(xs[c]+20,ys[r]+48),cv2.FONT_HERSHEY_SIMPLEX,1,(0,0,0),2)
        class Reader:
            def text_rec(self,crops):
                return [(s,.999) for s in ['1','1','2','2','3','3','5','5']],None
        result=verify_cells(image,detect_tables(image),{'lines':[{'words':words}]},lambda t:t,Reader())
        self.assertEqual(result['cell_ocr_changes'],[])
        self.assertEqual(result['cell_ocr_checked'],4)
        self.assertEqual(result['cell_ocr_issues'][0]['text'],'5')
        self.assertIn('quantity times price',result['cell_ocr_issues'][0]['reason'])


if __name__=='__main__':unittest.main()
