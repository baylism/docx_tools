import zipfile
from functions import parse, parse_notes

# Define XML document schema tags
TEXT_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = TEXT_NAMESPACE + 'p'
FOOTNOTE = TEXT_NAMESPACE + 'footnote'
ENDNOTE = TEXT_NAMESPACE + 'endnote'
TABLE = TEXT_NAMESPACE + 'tbl'

class Document:
    def __init__(self, docx_file):

        self.file_dir = docx_file

        # Unzip docx file 
        unzip = zipfile.ZipFile(docx_file, 'r')
        document_xml = unzip.read('word/document.xml')
        footnotes_xml = unzip.read('word/footnotes.xml')
        endnotes_xml = unzip.read('word/endnotes.xml')

        # Extract all XML files (for testing)
        unzip.extractall(path='docx_extract')
        unzip.close()

        # Ensure main document text in unicode 
        if not isinstance(document_xml, str):
            document_xml = document_xml.decode()

        # Parse XML files
        self.paragraphs = parse(document_xml, PARA)
        self.footnotes = parse_notes(footnotes_xml, FOOTNOTE)
        self.endnotes = parse_notes(endnotes_xml, ENDNOTE)
        self.tables = parse(document_xml, TABLE)

    def get_text(self):
        """ Returns Document text in a single string"""

        return '\n\n'.join(self.paragraphs)
    
    def get_paras(self):
        """ Returns list of paragraphs in Document 
        
        Returns None for empty Document
        """
        
        if self.paragraphs:
            return self.paragraphs
        
        return

    def get_footnotes(self):
        '''Returns a list of footnotes'''
        
        return self.footnotes
    def get_endnotes(self):
        '''Returns a list of endnotes'''
        
        return self.endnotes    
    
    def get_footnote_text(self):
        '''Returns footnote text in a single string'''

        try:
            return '\n\n'.join(self.footnotes)
        
        except TypeError:
            print('No footnotes in {0}'.format(self.file_dir))
            return

    def get_endnote_text(self):
        '''Returns footnote text in a single string'''

        try:
            return '\n\n'.join(self.endnotes)
        
        except TypeError:
            print('No endnotes in {0}'.format(self.file_dir))
            return
    def get_table_text(self):
        '''Returns footnote text in a single string'''

        return '\n\n'.join(self.tables)

    def count_words(self):
        """Returns word count"""
        return len(self.get_text().split())

    def count_chars(self):
        """Returns count of characters 
        includes: spaces, footnote/endnote placeholders, tabs, breaks
        
        returns 0 if no characters or spaces
        """
        
        note_placeholders = 0
        initial_tabs = 3

        if self.footnotes:
            note_placeholders = len(self.footnotes)

        return len(''.join(self.paragraphs)) + note_placeholders - initial_tabs

    def count_fn_chars(self):
        """
        Returns count of footnote characters 
        includes spaces

        Returns 0 if no footnotes
        """


        if self.footnotes:
            placeholder_chars = 0
            
            if len(self.footnotes) > 9:
                placeholder_chars = len(self.footnotes) - 9
            
            return len(''.join(self.footnotes)) - placeholder_chars
        
        return 0

    def count_en_chars(self):
        """
        Returns count of footnote characters 
        includes spaces

        returns 0 if no endnotes
        """
                
        if self.endnotes:
            placeholder_chars = 0
            
            if len(self.endnotes) > 9:
                placeholder_chars = len(self.endnotes) - 9
            
            return len(''.join(self.endnotes)) - placeholder_chars

        return 0
    
    def count_total_chars(self):
        """Returns total character count
        
        includes text, endnotes and footnotes
        """
        return self.count_chars() + self.count_en_chars() + self.count_fn_chars()

    def count_tables(self):
        if self.tables:
            return len(self.tables)

        else: 
            return 0

# Testing
#t = Document('docx_extract/testdoc1.docx')

#print(t.get_text())
#print(t.get_table_text())
#print(t.count_tables())

#print(t.get_footnotes())
#print(t.count_fn_chars())
#print(t.count_chars())
#print(t.get_paras())


#print(t.get_footnote_text())


#print(t.get_endnote_text())
#print(t.count_en_chars() + t.count_chars())




# Print info 
#print('Words: {0}'.format(words))
#print('Chars incl spaces: {0}'.format(chars))
#print('Chars no spaces: {0}'.format(chars_no_spaces))

#if eq_lines:
#    print('EQ lines: {0}'.format(eq_lines))

#print(paras)
# Print text to stdout
#p_output(output)

# Save text to File
#save_output(output)
