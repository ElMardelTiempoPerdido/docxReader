import os
from xml.etree.ElementTree import XML

import regex as re
from docx import Document


class DocxReader:
    """word reading cls"""

    def __init__(self, fpath):
        self.fpath = fpath
        self.fbasename = os.path.basename(fpath)
        self.doc = Document(fpath)
        self.fullcontent = ""

    @staticmethod
    def get_accepted_text(p, remove_strike=True):
        """
        get content from a docx paragraph, with all changes accepted

        :param p:                docx paragraph
        :param remove_strike:    True- addtionaly remove common strikethrough styled(not deleted in changes) content
        :return:
            full content with changes accepted
            is the end of this paragraph deleted(in changes)
        -------------------------------------------------------
        reference
        https://stackoverflow.com/questions/47666978/accepting-all-changes-in-a-ms-word-document-by-using-python
        """
        WORD_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        TEXT = WORD_NAMESPACE + "t"
        xml = p._p.xml

        # remove common strikethrough styled(different from being deleted in changes) content
        if remove_strike:
            pat = "<w:r w:r.*?strike.*?</w:r>"
            rev_pat = '>r:w/<.*?ekirts.*? r:w<'
            rev_xml = xml[::-1]
            repl = re.findall(rev_pat, rev_xml, re.DOTALL)
            repl = [x[::-1] for x in repl]
            dst_repl = []
            for rep in repl:
                dst_repl += re.findall(pat, rep, re.DOTALL)
            for rep in dst_repl:
                xml = xml.replace(rep, "")

        # remove <moveFrom>s to read moved texts for only once
        xml = re.sub("<w:moveFrom.*?/>", "", xml)
        xml = re.sub("<w:moveFrom.*?</w:moveFrom>", "", xml, flags=re.S)

        # soft enter to enter
        xml = xml.replace('<w:br/>', """<w:t>\n</w:t>""")

        # (maybe..)a single-line <w:del .../> indicates the end of this paragraph is deleted(in changes)
        del_enter = False
        if re.findall("<w:del .*?/>", xml):
            del_enter = True

        # get all <w:t> texts
        tree = XML(xml)
        runs = [node.text for node in tree.iter(TEXT) if node.text]

        return "".join(runs), del_enter

    def get_fullcontent(self, remove_strike=True):
        """
        get docx full content

        :param remove_strike:    True- addtionaly remove common strikethrough(not changes) content
        """
        self.fullcontent = ""
        last_force_concat_end = False  # is the enter of pre paragraph deleted
        for para in self.doc.paragraphs:
            text, force_concat_end = self.get_accepted_text(para, remove_strike)
            # concat 2 paragraphs without \n
            if last_force_concat_end:
                self.fullcontent += text
            else:
                self.fullcontent += "\n" + text
            last_force_concat_end = force_concat_end
            pass
        # remove the first \n
        self.fullcontent = self.fullcontent.lstrip("\n")

        return self.fullcontent


if __name__ == '__main__':
    d = DocxReader(os.path.join('sample.docx'))
    print(r"[remove strikethrough text]")
    print(d.get_fullcontent(remove_strike=True))
    print()
    print(r"[reserve strikethrough text]")
    print(d.get_fullcontent(remove_strike=False))
