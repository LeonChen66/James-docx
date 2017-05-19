import docx
import os

docx_name = input('Please input File Name ：')
docx_name = docx_name + '.docx'
print(docx_name)
document = docx.Document('example.docx')


table = document.tables

def auto_word_cell(replaced_word,content):
    for row in table[0].rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # if paragraph.text =="aa":
                    #     print(paragraph.text)
                    #     paragraph.text = "哈哈哈"
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if inline[i].text == replaced_word:
                            inline[i].text = content
                        # print(inline[i].text)
                        # inline[i].bold = True

# row = table[0].rows[0]

#
# cell = row.cells[6]
#
# paragraph = cell.paragraphs[0]
# print(paragraph.text)

auto_word_cell('aa','hahaha')
document.save(docx_name)
os.system('pause')