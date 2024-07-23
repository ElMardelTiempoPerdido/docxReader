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
        :param remove_strike:    True- addtionaly remove common strikethrough(not changes) content
        :return:
            full content with changes accepted
            is the beginning of this paragraph deleted(in changes)
            is the end of this paragraph deleted(in changes)
        -------------------------------------------------------
        reference
        https://stackoverflow.com/questions/47666978/accepting-all-changes-in-a-ms-word-document-by-using-python
        """
        start_with_del, end_with_del = False, False
        WORD_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        TEXT = WORD_NAMESPACE + "t"
        xml = p._p.xml

        # remove common strikethrough content
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

        # index of the first w:delText tag < index of the first w:t tag
        # -> the beginning of this paragraph is deleted(in changes)
        first_del_index = xml.find("<w:delText")
        first_t_index = xml.find('<w:t')
        if first_del_index != -1 and first_del_index < first_t_index:
            start_with_del = True

        # index of the last w:delText tag > index of the last w:t tag
        # -> the end of this paragraph is deleted(in changes)
        last_del_index = xml.rfind("<w:delText")
        last_t_index = xml.rfind('<w:t')
        if last_del_index > last_t_index:
            end_with_del = True

        # get all text from w:t tag
        tree = XML(xml)
        runs = [node.text for node in tree.iter(TEXT) if node.text]
        return "".join(runs), start_with_del, end_with_del

    def get_fullcontent(self, remove_strike=True, is_concat=True):
        """
        get docx full content

        :param remove_strike:    True- addtionaly remove common strikethrough(not changes) content
        :param is_concat:        True- concat two paragraphs before and after without \n
                                if threre is a delete change across two lines
        """
        self.fullcontent = ''
        last_start_with_del, last_end_with_del = False, False
        for para in self.doc.paragraphs:
            text, start_with_del, end_with_del = self.get_accepted_text(para, remove_strike)
            # concat 2 paragraphs without \n
            if start_with_del and last_end_with_del and is_concat:
                self.fullcontent += text
            else:
                self.fullcontent += "\n" + text
            last_start_with_del, last_end_with_del = start_with_del, end_with_del
            pass
        # remove the first \n
        self.fullcontent = self.fullcontent.lstrip("\n")
        return self.fullcontent


if __name__ == '__main__':
    d = DocxReader(os.path.join('sample.docx'))
    print(r"[concat without \n]")
    print(d.get_fullcontent(remove_strike=True, is_concat=True))
    print()
    print(r"[concat without \n + reserve strikethrough text]")
    print(d.get_fullcontent(remove_strike=False, is_concat=True))
    print()
    print(r"[concat with \n]")
    print(d.get_fullcontent(remove_strike=True, is_concat=False))
