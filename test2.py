import docx


document = docx.Document('james.docx')


table = document.tables

for row in table[0].rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            # if paragraph.text =="aa":
            #     print(paragraph.text)
            #     paragraph.text = "哈哈哈"
            inline = paragraph.runs
            for i in range(len(inline)):
                if inline[i].text == "陳":
                    inline[i].text = "asdjfl"
                # print(inline[i].text)
                # inline[i].bold = True

# row = table[0].rows[0]

#
# cell = row.cells[6]
#
# paragraph = cell.paragraphs[0]
print(paragraph.text)
document.save('test.docx')