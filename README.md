# docx_tools

## Overview
Functions for extracting text from DOCX files and performing rough word and character counts. 

Extracted text can be stored as lists of sentences or paragraphs, or saved to a text file. 

## Important Note
Complex characters are replaced with their ASCII symbol names. 

For basic document elements (paragraphs, footnotes, formulae) the counting methods have been designed to mirror the counts provided in Microsoft Word. Despite this, there is likely to be some variation for most documents as a number of text elements (including hyperlinks and diagrams) are not yet supported. 
