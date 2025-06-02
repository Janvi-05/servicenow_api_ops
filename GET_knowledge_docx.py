import requests
import json
from bs4 import BeautifulSoup
from datetime import datetime
import re
import os
import argparse
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

def format_kb_article_backup(article):
    """Format a single knowledge base article for text backup"""
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

def format_kb_article_to_docx(doc, article):
    """Add a formatted knowledge base article to the Word document"""
    
    # Article title/number
    if article.get('number'):
        title = doc.add_heading(f"Article: {article['number']}", level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Metadata table
    meta_fields = [
        ("Created:", article.get('sys_created_on')),
        ("Updated:", article.get('sys_updated_on')),
        ("Status:", article.get('workflow_state')),
        ("Published:", article.get('published')),
        ("Author:", article.get('author', {}).get('display_value') if isinstance(article.get('author'), dict) else article.get('author')),
        ("View Count (All):", article.get('u_view_count_all')),
        ("Active:", article.get('active')),
        ("Topic:", article.get('topic')),
        ("Valid To:", article.get('valid_to')),
        ("KB Category:", article.get('kb_category', {}).get('display_value') if isinstance(article.get('kb_category'), dict) else article.get('kb_category')),
    ]

    # Filter out None values and create table
    meta_fields = [item for item in meta_fields if item[1]]
    if meta_fields:
        table = doc.add_table(rows=0, cols=2)
        table.style = 'Table Grid'
        for label, value in meta_fields:
            row = table.add_row()
            row.cells[0].text = label
            row.cells[1].text = value

        # Add space after table
        doc.add_paragraph()

    # Main content
    if article.get('text'):
        content_heading = doc.add_heading('Content', level=2)
        clean_text = clean_html_text(article['text'])
        
        # Split content into paragraphs and add them
        paragraphs = clean_text.split('\n\n')
        for para_text in paragraphs:
            if para_text.strip():
                doc.add_paragraph(para_text.strip())
    
    # Add page break between articles (except for the last one)
    doc.add_page_break()

def download_attachments_for_article(table_sys_id, output_dir, headers):
    """Download attachments for a specific KB article and save them in its folder"""
    attachment_url = f"https://lendlease.service-now.com/api/now/attachment?sysparm_query=table_sys_id={table_sys_id}"
    
    try:
        response = requests.get(attachment_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            attachments = data.get('result', [])
            
            if not attachments:
                print(f"📎 No attachments found for {table_sys_id}")
                return
            
            print(f"📎 Found {len(attachments)} attachment(s) for {table_sys_id}")
            
            for attachment in attachments:
                file_name = attachment.get('file_name')
                download_link = attachment.get('download_link')
                file_size = attachment.get('size_bytes')
                
                if download_link and file_name:
                    try:
                        file_response = requests.get(download_link, headers=headers)
                        if file_response.status_code == 200:
                            file_path = os.path.join(output_dir, file_name)
                            with open(file_path, 'wb') as f:
                                f.write(file_response.content)
                            print(f"   ✓ Downloaded: {file_name} ({file_size} bytes)")
                        else:
                            print(f"   ✗ Failed to download {file_name} (Status {file_response.status_code})")
                    except Exception as e:
                        print(f"   ✗ Error downloading {file_name}: {e}")
        else:
            print(f"❌ Failed to get attachment list for {table_sys_id}. Status code: {response.status_code}")
    except Exception as e:
        print(f"❌ Exception while fetching attachments: {e}")


# Parse command-line argument for knowledge base ID
parser = argparse.ArgumentParser(description='Download and export KB articles from ServiceNow')
parser.add_argument('kb_id', type=str, help='Knowledge Base sys_id (e.g., 01125e5a1b9b685017eeebd22a4bcb44)')
args = parser.parse_args()
kb_id = args.kb_id


# Your API call
url = f"https://lendlease.service-now.com/api/now/table/kb_knowledge?sysparm_query=sys_class_name!=^publishedISNOTEMPTY^latest=true^kb_knowledge_base={kb_id}&sysparm_display_value=true"
payload = {}
headers = {
  'Authorization': 'Bearer LFT612CwbhRj2QR_id2LcEojIWraGE9lgpu51BabMHszqcWclNCQIpqHSaDO-pmwGXBowMINr1d_ENTsypdIeQ',
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
        
        # Create Word document
        doc = Document()
        
        # Document title and header
        title = doc.add_heading('Lendlease Knowledge Base Articles', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Document info
        info_para = doc.add_paragraph()
        info_para.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        info_para.add_run(f"Total Articles: {len(data.get('result', []))}")
        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add separator
        doc.add_paragraph("_" * 80)
        
        # Process each article
        articles = data.get('result', [])
        for i, article in enumerate(articles):
            # Generate a new document for each article
            doc = Document()

            # Document title
            title = doc.add_heading('Lendlease Knowledge Base Article', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Article info
            info_para = doc.add_paragraph()
            info_para.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            info_para.add_run(f"Article Number: {article.get('number', 'Unknown')}")
            info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph("_" * 80)

            # Add article content
            format_kb_article_to_docx(doc, article)

            # Save each article as a separate .docx file
            safe_article_number = re.sub(r'[^\w\-_. ]', '_', article.get('number', f"article_{i+1}"))
            # Create output directory if it doesn't exist
            # output_dir = "KB_docx_files_{timestamp}".format(timestamp=timestamp)
            # os.makedirs(output_dir, exist_ok=True)
            # Create parent directory
            parent_dir = f"KB_docx_files_{timestamp}"
            os.makedirs(parent_dir, exist_ok=True)

            # Create subfolder for each article number
            article_number = article.get('number', f"article_{i+1}")
            safe_article_number = re.sub(r'[^\w\-_. ]', '_', article_number)
            article_dir = os.path.join(parent_dir, safe_article_number)
            os.makedirs(article_dir, exist_ok=True)


            # Save the .docx file into the folder
            docx_filename = f"kb_article_{safe_article_number}.docx"
            docx_path = os.path.join(article_dir, docx_filename)
            doc.save(docx_path)
            print(f"📄 Saved: {docx_path}")
            
            # Download attachments for this article
            table_sys_id = article.get('sys_id')
            if table_sys_id:
                download_attachments_for_article(table_sys_id, article_dir, headers)
            else:
                print(f"⚠️ No sys_id found for article {safe_article_number}, skipping attachment download.")

        
       
        print(f"📊 Processed {len(articles)} articles")
        
       
        
    else:
        print(f"❌ API request failed with status code: {response.status_code}")
        print(f"Response: {response.text[:200]}...")

except requests.exceptions.RequestException as e:
    print(f"❌ Request failed: {e}")
except json.JSONDecodeError as e:
    print(f"❌ Failed to parse JSON response: {e}")
except Exception as e:
    print(f"❌ An error occurred: {e}")

