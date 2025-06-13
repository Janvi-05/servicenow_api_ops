from bs4 import BeautifulSoup
import html

def extract_text_from_spans(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    
    # Get all text inside spans, concatenated together
    text = "".join(span.get_text() for span in soup.find_all("span"))
    
    # Unescape HTML entities (&lsquo;, &rsquo;, etc)
    clean_text = html.unescape(text)
    
    # Strip leading/trailing whitespace and replace multiple spaces/newlines with single space
    clean_text = " ".join(clean_text.split())
    
    return clean_text

html_snippet = '''
<span xml:lang="EN-AU" data-contrast="auto">The manager can click on the Related Action Item &lsquo;ie. </span>
<span xml:lang="EN-AU" data-contrast="auto">3 Dots</span>
<span xml:lang="EN-AU" data-contrast="auto">&rsquo; go to &lsquo;</span>
<span xml:lang="EN-AU" data-contrast="auto">Job Change</span>
<span xml:lang="EN-AU" data-contrast="auto">&rsquo; and select &lsquo;</span>
<span xml:lang="EN-AU" data-contrast="auto">Edit Job Requisition</span>
<span xml:lang="EN-AU" data-contrast="auto">&rsquo;.</span>
'''

print(extract_text_from_spans(html_snippet))
