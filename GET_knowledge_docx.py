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
        ("Updated By:", article.get('sys_updated_by')),
        ("Sys Domain:", article.get('sys_domain', {}).get('display_value') if isinstance(article.get('sys_domain'), dict) else article.get('sys_domain')),
        ("x_caukp_ebonding_no_return:", article.get('x_caukp_ebonding_no_return')),
        ("Select Portal:", article.get('u_select_portal', {}).get('display_value') if isinstance(article.get('u_select_portal'), dict) else article.get('u_select_portal')),
        ("Status:", article.get('workflow_state')),
        ("Created By:", article.get('sys_created_by')),
        ("Published:", article.get('published')),
        ("Author:", article.get('author', {}).get('display_value') if isinstance(article.get('author'), dict) else article.get('author')),
        ("x_caukp_ebonding_integration_mode:", article.get('x_caukp_ebonding_integration_mode')),
        ("Helpful Count:", article.get('helpful_count')),
        ("System Domain Path:", article.get('sys_domain_path')),
        ("View Count (All):", article.get('u_view_count_all')),
        ("Version:", article.get('version', {}).get('display_value') if isinstance(article.get('version'), dict) else article.get('version')),
        ("Active:", article.get('active')),
        ("Topic:", article.get('topic')),
        ("Valid To:", article.get('valid_to')),
        ("KB Category:", article.get('kb_category', {}).get('display_value') if isinstance(article.get('kb_category'), dict) else article.get('kb_category')),
        ("Meta description:",article.get('meta_description')),
        ("KB Knowledge base:", article.get('kb_knowledge_base', {}).get('display_value') if isinstance(article.get('kb_knowledge_base'), dict) else article.get('kb_knowledge_base')),
        ("Meta:",article.get('meta')),
        ("U Problem Type:",article.get('u_problem_type')),
        ("Display number:",article.get('display_number')),
        ("Base version:",article.get('base_version',{}).get('display_value') if isinstance(article.get('base_version'), dict) else article.get('base_version')),
        ("Short description:",article.get('short_description')),
        ("Direct:",article.get('direct')),
        ("Disable suggesting:",article.get('disable_suggesting')),
        ("Class name:",article.get('sys_class_name')),
        ("Article Id:",article.get('article_id')),
        ("Sys Id:",article.get('sys_id')),
        ("Use Count:",article.get('use_count')),
        ("Flagged:",article.get('flagged')),
        ("Disable commenting:",article.get('disable_commenting')),
        ("Adding to home page:",article.get('u_add_to_homepage')),
        ("Display attachments:",article.get('display_attachments')),
        ("Latest:",article.get('latest')),
        ("Summary:",article.get('summary',{}).get('display_value') if isinstance(article.get('summary'), dict) else article.get('summary')),
        ("Sys View Count:",article.get('sys_view_count')),
        ("Revised by:",article.get('revised_by',{}).get('display_value') if isinstance(article.get('revised_by'), dict) else article.get('revised_by')),
        ("Article Type:",article.get('article_type')),
        ("Needs review:",article.get('u_needs_review')),
        ("Sys Mod Count:",article.get('sys_mod_count')),
        ("View as allowed:",article.get('view_as_allowed')),
        ("Category:",article.get('category')),
        ("Reminder send date:",article.get('u_reminder_send_date')),
        ("Wiki:",article.get('wiki')),
        ("Rating:",article.get('rating"')),
        ("Source:",article.get('source')),
        ("x_caukp_ebonding_sdc:",article.get('x_caukp_ebonding_sdc')),
        ("Scheduled Publish date:",article.get('scheduled_publish_date')),
        ("Image:",article.get('image')),
        ("KBI Uniqueid",article.get('u_kbi_uniqueid')),
        ("cmdb ci:",article.get('cmdb_ci')),
        ("Can Read User Criteria:",article.get('can_read_user_criteria')),
        ("Cannot Read User Criteria:",article.get('cannot_read_user_criteria')),
        ("x caukp ebonding requester id:",article.get('x_caukp_ebonding_requester_id')),
        ("Last Review date:",article.get('u_last_review_date')),
        ("x caukp ebonding provider id:",article.get('x_caukp_ebonding_provider_id')),
        ("Roles:",article.get('roles')),
        ("Description:",article.get('description')),
        ("sn_grc_target_table:",article.get('sn_grc_target_table')),
        ("Retired:",article.get('retired')),
        ("Video URL:",article.get('u_video_url')),
        ("sn_grc_source:",article.get('sn_grc_source')),
        ("Sys Tags:",article.get('sys_tags')),
        ("Replacement Article:",article.get('replacement_article')),
        ("x caukp ebonding provider:",article.get('x_caukp_ebonding_provider')),
        ("Taxonomy Topic:",article.get('taxonomy_topic')),
        ("x caukp ebonding requester:",article.get('x_caukp_ebonding_requester')),
        ("Ownership group:",article.get('ownership_group')),
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


# Parse command-line argument for knowledge base ID
parser = argparse.ArgumentParser(description='Download and export KB articles from ServiceNow')
parser.add_argument('kb_id', type=str, help='Knowledge Base sys_id (e.g., 01125e5a1b9b685017eeebd22a4bcb44)')
args = parser.parse_args()
kb_id = args.kb_id


# Your API call
url = f"https://lendlease.service-now.com/api/now/table/kb_knowledge?sysparm_query=sys_class_name!=^publishedISNOTEMPTY^latest=true^kb_knowledge_base={kb_id}&sysparm_display_value=true"
payload = {}
headers = {
  'Authorization': 'Bearer T8cdGUPzDMQGwpGEc5zUYTrqf46fCmW9HQgO58i9ZTOpk_c03NWthebCKEMcjMy-GF50Ipo7Bt3bpydLFgTtNA',
  'Cookie': 'BIGipServerpool_lendlease=79ce09d0b3e23f5256871ffb409399ee; JSESSIONID=01A4CE5F22AFEE5A3996E5B5D207CF7F; glide_node_id_for_js=183c289393367ee4c6986408359e1a4703ca618e44d342a23d2196f17ea03570; glide_session_store=90DEDC7D2B352A507DDDFD84FE91BF59; glide_user_activity=U0N2M18xOlFHZ0dwM0RGdmVkNUJpaXFaUURYRmRqeTNuQmNDdTNkVnFvbXk1dDE4QTQ9OjJFc0pxd2x0eEdBdGlWNmRZaDdYSVowc29tQ2RHeVFaenJNVGpMM0ptcGM9; glide_user_route=glide.da3ebff2633845391cdfe258b9a97c5d'
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
            output_dir = "KB_docx_files_{timestamp}".format(timestamp=timestamp)
            os.makedirs(output_dir, exist_ok=True)

            # Save the .docx file into the folder
            docx_filename = f"kb_article_{safe_article_number}_{timestamp}.docx"
            docx_path = os.path.join(output_dir, docx_filename)
            doc.save(docx_path)
            print(f"📄 Saved: {docx_path}")

        
       
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

print("\n🔧 Required libraries: pip install requests beautifulsoup4 python-docx")