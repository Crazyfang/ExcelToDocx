from docxcompose.composer import Composer
from docx import Document


def main(files,final_docx):
    new_document = Document()
    composer = Composer(new_document)
    for fn in files:
        composer.append(Document(fn))
    composer.save(final_docx)


main(['/Users/fangyong/Desktop/123.docx', '/Users/fangyong/Desktop/321.docx'], '/Users/fangyong/Desktop/213.docx')