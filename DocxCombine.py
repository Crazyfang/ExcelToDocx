from docxcompose.composer import Composer
from docx import Document
import os


def main(files, final_docx=None):
    if final_docx:
        pass
    else:
        final_docx = os.path.join(os.path.dirname(files[0]), '合并.docx')

    try:
        new_document = Document()
        composer = Composer(new_document)
        for fn in files:
            composer.append(Document(fn))
        composer.save(final_docx)

        return [True, '']
    except Exception as e:
        return [False, e]

if __name__ == '__main__':
    main(['/Users/fangyong/Desktop/123.docx', '/Users/fangyong/Desktop/321.docx'], '/Users/fangyong/Desktop/213.docx')