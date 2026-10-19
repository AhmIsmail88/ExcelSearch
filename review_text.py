"""Bounded, plain-text word diffs; HTML escaping remains the UI's responsibility."""
import re
from difflib import SequenceMatcher


def segments(left,right):
    left='' if left is None else str(left)
    right='' if right is None else str(right)
    if left==right:return [{'text':left,'changed':False}],[{'text':right,'changed':False}]
    if len(left)+len(right)>24000:
        return [{'text':left,'changed':True}],[{'text':right,'changed':True}]
    a,b=re.findall(r'\s+|\S+',left),re.findall(r'\s+|\S+',right)
    if len(a)+len(b)>4000:
        return [{'text':left,'changed':True}],[{'text':right,'changed':True}]
    out_a,out_b=[],[]
    for op,i,j,k,l in SequenceMatcher(None,a,b).get_opcodes():
        if i<j:out_a.append({'text':''.join(a[i:j]),'changed':op!='equal'})
        if k<l:out_b.append({'text':''.join(b[k:l]),'changed':op!='equal'})
    return out_a,out_b
