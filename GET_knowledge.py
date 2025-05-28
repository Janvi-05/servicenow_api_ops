import requests
import json
from bs4 import BeautifulSoup
from datetime import datetime
import re

def clean_html_text(html_content):
    """Convert HTML to clean text"""
    if not html_content:
        return ""
    
    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Convert to text while preserving some structure
    text = soup.get_text()
    
    # Clean up extra whitespace and line breaks
    text = re.sub(r'\n\s*\n', '\n\n', text)  # Multiple newlines to double newline
    text = re.sub(r'[ \t]+', ' ', text)       # Multiple spaces/tabs to single space
    text = text.strip()
    
    return text

def format_kb_article(article):
    """Format a single knowledge base article"""
    formatted = []
    
    # Title/Number
    if article.get('number'):
        formatted.append(f"Article: {article['number']}")
        formatted.append("=" * 50)
    
    # Metadata
    if article.get('sys_created_on'):
        formatted.append(f"Created: {article['sys_created_on']}")
    if article.get('sys_updated_on'):
        formatted.append(f"Updated: {article['sys_updated_on']}")
    if article.get('workflow_state'):
        formatted.append(f"Status: {article['workflow_state']}")
    
    formatted.append("")  # Empty line
    
    # Main content
    if article.get('text'):
        clean_text = clean_html_text(article['text'])
        formatted.append("CONTENT:")
        formatted.append("-" * 20)
        formatted.append(clean_text)
    
    formatted.append("\n" + "="*80 + "\n")  # Separator between articles
    
    return "\n".join(formatted)

url = "https://lendlease.service-now.com/api/now/table/kb_knowledge?sysparm_query=sys_class_name!=^publishedISNOTEMPTY^latest=true^kb_knowledge_base=01125e5a1b9b685017eeebd22a4bcb44&sysparm_display_value=true"

payload = {}
headers = {
  'Authorization': 'Bearer ApUB-mr5xOiFMChMMopPKM5EgXbEHpNj8rbxNiZ62gVfVeWXCTueytgPk0IydGFvc1OdDFv4GTIvvhC69wkX7g',
  'Cookie': 'BIGipServerpool_lendlease=c5889ad29f701618e3baa37002034b82; JSESSIONID=3901AC59B602B51CE1CF74C8956FD362; glide_node_id_for_js=fc4812175032dd94c0ff92cf846b17cf27f0dce0a6beb49e12e5c7bb0f48d836; glide_session_store=6360D6592B3D6E50E412F41CD891BF5D; glide_user_activity=U0N2M18xOnRMdkppdFlTN2o2cFlnUVdaQ092UjZ6S0pFdXV0dmZBb3BMcGxVa0hrZ1E9OlVBQWc4QWozUERYQi9mVCs2WDRJa0hTRTgwQjkxMGZkMzUrNGxlUXRNUW89; glide_user_route=glide.5a07cc0a1b859ed021434a69d48daaeb'
}


try:
    # Make API request
    response = requests.request("GET", url, headers=headers, data=payload)
    
    if response.status_code == 200:
        # Parse JSON response
        data = response.json()
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create document content
        document_content = []
        document_content.append("LENDLEASE KNOWLEDGE BASE ARTICLES")
        document_content.append("=" * 50)
        document_content.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        document_content.append(f"Total Articles: {len(data.get('result', []))}")
        document_content.append("\n" + "="*80 + "\n")
        
        # Process each article
        for article in data.get('result', []):
            formatted_article = format_kb_article(article)
            document_content.append(formatted_article)
        
        # Save to text file
        txt_filename = f"lendlease_kb_articles_{timestamp}.txt"
        with open(txt_filename, 'w', encoding='utf-8') as file:
            file.write("\n".join(document_content))
        
        print(f"✅ Document saved as: {txt_filename}")
        print(f"📊 Processed {len(data.get('result', []))} articles")
        
        # Optional: Also save as markdown for better formatting
        md_filename = f"lendlease_kb_articles_{timestamp}.md"
        with open(md_filename, 'w', encoding='utf-8') as file:
            # Convert to markdown format
            md_content = "\n".join(document_content)
            md_content = md_content.replace("=" * 50, "---")
            md_content = md_content.replace("=" * 80, "\n---\n")
            file.write(f"# Lendlease Knowledge Base Articles\n\n{md_content}")
        
        print(f"✅ Markdown version saved as: {md_filename}")
        
    else:
        print(f"❌ API request failed with status code: {response.status_code}")
        print(f"Response: {response.text[:200]}...")

except requests.exceptions.RequestException as e:
    print(f"❌ Request failed: {e}")
except json.JSONDecodeError as e:
    print(f"❌ Failed to parse JSON response: {e}")
except Exception as e:
    print(f"❌ An error occurred: {e}")

print("\n🔧 Required libraries: pip install requests beautifulsoup4")