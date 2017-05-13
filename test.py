# -*- coding: utf-8 -*-
"""
Created on Sat May  6 02:43:07 2017

@author: Leon
"""
from docx import Document

def replace_string(filename):
    doc = Document(filename)
    for p in doc.paragraphs:
        if '助理學經歷表' in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                print(inline[i].text)
                if '助理學經歷表' in inline[i].text:
                    text = inline[i].text.replace('助理學經歷表', 'new text')
                    inline[i].text = text
            print(p.text)

    doc.save('dest1.docx')
    return 1

replace_string('研究助理學經歷表.docx')