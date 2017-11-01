import xml.etree.cElementTree as ET

# DOCX namespace
TEXT_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TEXT = TEXT_NAMESPACE + 't'
PARA = TEXT_NAMESPACE + 'p'
FOOTNOTE = TEXT_NAMESPACE + 'footnote'
ENDNOTE = TEXT_NAMESPACE + 'endnote'
NOTE_ID = TEXT_NAMESPACE + 'id'
TAB_ROW = TEXT_NAMESPACE + 'tr'
TABLE = TEXT_NAMESPACE + 'tbl'

EQUATION_NAMESPACE = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
EQ_PARA = EQUATION_NAMESPACE + 'oMathPara'
EQ_LINE = EQUATION_NAMESPACE + 'oMath'
EQ_TEXT = EQUATION_NAMESPACE + 't'

FORMATTING_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
BREAK = FORMATTING_NAMESPACE + 'br'
TAB = FORMATTING_NAMESPACE + 'tab'


def save_output(output):
    """ Save output to file 'extracted text.txt'. 
    
        If text contains unusual symbols, write raw text 
        and replace with symbol names.
        
    """

    try:
        text_file = open('extracted text.txt', 'w')
        text_file.write(output)
        text_file.close()

    except UnicodeEncodeError:
        text_file = open('extracted text.txt', 'wb')
        text_file.write(output.encode('ascii', errors='namereplace'))
        text_file.close()

def p_output(output):


    """ Print output to stdout. 

    If text contains unusual symbols, print raw text 
    and replace symbol names.
    
    """
    try:
        print(output)
    except UnicodeEncodeError:
        print(output.encode('ascii', errors='namereplace'))

def parse1(text_string, element_type):
    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags    
    """
    tree = ET.fromstring(text_string)
    elements = []

    for paras in tree.iter(element_type):
        para_text = []

        for child in paras.iter():
            if child.tag in (TEXT, EQ_TEXT):
                if child.text:
                    para_text.append(child.text)
            
            elif child.tag == BREAK:
                para_text.append('\n')
            
            elif child.tag == TAB:
                para_text.append('\t')
        if para_text:
            elements.append(''.join(para_text))
    
    return elements

def parse(text_string, element_type):
    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags    
    """
    root = ET.fromstring(text_string)

  
    list_of_paragraphs = []

    for paragraph in root[0].findall(element_type):
        para_text = []

        for item in paragraph.iter():
        
            if item.tag in (TEXT, EQ_TEXT):
                if item.text:
                    para_text.append(item.text)
            
            elif item.tag == BREAK:
                para_text.append('\n')
            
            elif item.tag == TAB:
                para_text.append('\t')
            
        if para_text:
            list_of_paragraphs.append(''.join(para_text))

    return list_of_paragraphs

def parse_fn(text_string):
    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags    
    """
    tree = ET.fromstring(text_string)
    elements = []

    for footnote in tree.iter(FOOTNOTE):
        fn_paras = []

        for paras in footnote.iter(PARA):
            para_text = []

            for child in paras.iter():
                if child.tag in (TEXT, EQ_TEXT):
                    if child.text:
                        para_text.append(child.text)
                
                elif child.tag == BREAK:
                    para_text.append('\n')
                
                elif child.tag == TAB:
                    para_text.append('\t')

            fn_paras.append(''.join(para_text))
        
        elements.append(fn_paras)

    return elements[2:]

def parse_notes(text_string, element_type):

    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags 

        returns - a list of footnotes   
    """

    tree = ET.fromstring(text_string)
    elements = []

    for note in tree.iter(element_type):
        para_text = [note.attrib[NOTE_ID],]

        for child in note.iter():
            if child.tag in (TEXT, EQ_TEXT):
                if child.text:
                    para_text.append(child.text)

            elif child.tag == BREAK:
                para_text.append('\n')
            
            elif child.tag == TAB:
                para_text.append('\t')

        elements.append(''.join(para_text))

    if len(elements) > 2:    
        return elements[2:]
    else:
        return

def parse_tables1(text_string, element_type):
    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags 

        returns - a list of footnotes   
    """

    root = ET.fromstring(text_string)
    elements = {}

    for note in tree.iter(TAB_ROW):
        para_text = [note.attrib[NOTE_ID],]
 
        for child in note.iter():
            if child.tag in (TEXT, EQ_TEXT):
                if child.text:
                    para_text.append(child.text)

            elif child.tag == BREAK:
                para_text.append('\n')
            
            elif child.tag == TAB:
                para_text.append('\t')

        elements.append(''.join(para_text))

    if len(elements) > 2:    
        return elements[2:]
    else:
        return
def parse_tables(text_string, element_type):
    """ Parse xml file and return text of type 
        text_string - string of XML
        element_types - a tuple of XML namespace tags    
    """
    root = ET.fromstring(text_string)

  
    list_of_paragraphs = []

    for paragraph in root[0].findall(element_type):
        para_text = []

        for item in paragraph.iter():
        
            if item.tag in (TEXT, EQ_TEXT):
                if item.text:
                    para_text.append(item.text)
            
            elif item.tag == BREAK:
                para_text.append('\n')
            
            elif item.tag == TAB:
                para_text.append('\t')
            
        if para_text:
            list_of_paragraphs.append(''.join(para_text))
