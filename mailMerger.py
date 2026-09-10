from docx import Document

def mergeTXT(docx_path:str,strReplaceWhat:str,strReplaceWith:str):
    #create plain text version of email: open .docx file, replace [Anrede] for plain text-version of email
    text = []
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        #print(paragraph.text)
        if strReplaceWhat in paragraph.text:
            paragraph.text = strReplaceWith

        text.append(paragraph.text)
    
    text = '\n'.join(text)
    #print(text)
    return text

def mergeHTML(html_path:str,strReplaceWhat:str,strReplaceWith):
    #create html-version of email: open .htm file, replace [Anrede] for html-version of email
    html=[]
    with open(html_path,'r',encoding="utf-8",errors="ignore") as f:
        html = f.read()
        html = html.replace(strReplaceWhat,strReplaceWith)
    #print(html)
    return html
