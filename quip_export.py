import os
import sys
import json
import requests
import argparse
import quip
import time
import re
from urllib.parse import urlparse


def sanitize_filename(filename):
    """Remove invalid characters from a filename."""
    return re.sub(r'[\\/*?:"<>|]', '', filename)


def create_directory_if_not_exists(path):
    """Create a directory if it doesn't exist."""
    if not os.path.exists(path):
        os.makedirs(path)


def extract_id_from_link(link):
    """Extract the folder ID from a Quip link."""
    # Handle different Quip URL formats
    parsed_url = urlparse(link)
    path = parsed_url.path.strip('/')
    
    # Extract the ID - usually the last part of the path
    parts = path.split('/')
    if not parts:
        return None
        
    # The ID is typically the last segment in the URL
    potential_id = parts[-1]
    
    # Some Quip URLs include a hash fragment for document sections
    # Remove any hash part if present
    potential_id = potential_id.split('#')[0]
    
    return potential_id


def get_domain_from_link(link):
    """Extract the domain from a Quip link to determine the API base URL."""
    parsed_url = urlparse(link)
    return parsed_url.netloc


def download_blob(url, file_path):
    """Download an attachment to a file."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(file_path, 'wb') as f:
            f.write(response.content)
        return True
    except Exception as e:
        print(f"Error downloading attachment: {e}")
        return False


def process_thread(client, thread_id, current_path):
    """Process a document thread to extract content and attachments."""
    try:
        thread_data = client.get_thread(thread_id)
        
        if 'error' in thread_data:
            print(f"Error accessing document {thread_id}: {thread_data['error']}")
            return
        
        thread_title = thread_data.get('thread', {}).get('title', 'Unknown_Document')
        thread_title = sanitize_filename(thread_title)
        
        # Construct the export URL for DOCX format
        export_url = f"{client.base_url}/1/threads/{thread_id}/export/docx"
        headers = {"Authorization": f"Bearer {client.access_token}"}
        
        # Download the DOCX file
        docx_path = os.path.join(current_path, f"{thread_title}.docx")
        response = requests.get(export_url, headers=headers)
        
        if response.status_code == 200:
            with open(docx_path, 'wb') as file:
                file.write(response.content)
            print(f"Extracted DOCX: {docx_path}")
        else:
            print(f"Failed to download DOCX for {thread_title}: {response.status_code}")
            
            # Fallback to HTML if DOCX export fails
            html_path = os.path.join(current_path, f"{thread_title}.html")
            content = thread_data.get('html')
            if content:
                with open(html_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                print(f"Extracted HTML instead: {html_path}")
        # Download attachments if any
        blobs = thread_data.get('blobs', [])
        if blobs:
            # Create attachments folder
            attachments_path = os.path.join(current_path, f"{thread_title}_attachments")
            create_directory_if_not_exists(attachments_path)
            
            for blob in blobs:
                blob_url = blob.get('url')
                blob_name = blob.get('name', 'Unknown_Attachment')
                blob_name = sanitize_filename(blob_name)
                blob_path = os.path.join(attachments_path, blob_name)
                
                if blob_url and download_blob(blob_url, blob_path):
                    print(f"Downloaded attachment: {blob_path}")
    
    except Exception as e:
        print(f"Error processing document {thread_id}: {e}")


def process_folder(client, folder_id, base_path='.', depth=0):
    """Recursively process a Quip folder, extracting all files with hierarchical paths."""
    try:
        # Add indentation for better visual hierarchy in the console output
        indent = "  " * depth
        
        folder = client.get_folder(folder_id)
        
        if 'error' in folder:
            print(f"{indent}Error accessing folder {folder_id}: {folder['error']}")
            return
        print(f"DEBUG - Folder structure: {json.dumps(folder, indent=2)[:200]}...")
        folder_title = folder.get('folder', {}).get('title', 'Unknown_Folder')
        folder_title = sanitize_filename(folder_title)
        
        # Create a directory for the folder
        current_path = os.path.join(base_path, folder_title)
        create_directory_if_not_exists(current_path)
        print(f"{indent}Processing folder: {folder_title}")
        
        # Process all children in the folder
        children = folder.get('children', [])

        # Try different paths to find children
        children = []
        if 'children' in folder:
            children = folder['children']
        # elif 'folder' in folder and 'children' in folder['folder']:
        #     children = folder['folder']['children']
        # elif 'threads' in folder:
        #     children = folder['threads']
        
        # Count items for progress tracking
        total_items = len(children)
        print(f"Found in {folder_title}: ({total_items} items)")
        
        # Count items for progress tracking
        processed_items = 0
        
        for child in children:
            processed_items += 1
            
            # Determine type based on which ID is present
            if 'thread_id' in child:
                # Process documents
                print(f"{indent}  ({processed_items}/{total_items}) Processing document...")
                process_thread(client, child['thread_id'], current_path)
            elif 'folder_id' in child:
                # Recursively process subfolders
                process_folder(client, child['folder_id'], current_path, depth + 1)
            
            # Avoid hitting rate limits
            time.sleep(0.5)
            
        print(f"{indent}Completed folder: {folder_title} ({total_items} items)")
    
    except Exception as e:
        print(f"Error processing folder {folder_id}: {e}")


def main():
    parser = argparse.ArgumentParser(description='Extract files from a Quip folder with hierarchical path structure.')
    parser.add_argument('--token', '-t', dest='access_token', help='Your Quip access token')
    parser.add_argument('--folder', '-f', dest='folder_link', help='Link to the Quip folder')
    parser.add_argument('--output', '-o', default='./quip_export', help='Output directory (default: ./quip_export)')
    parser.add_argument('--token-file', dest='token_file', help='Path to file containing your Quip access token')
    parser.add_argument('--api-url', dest='api_url', help='Custom Quip API URL (default: https://platform.quip-amazon.com)')
    
    args = parser.parse_args()
    
    # Get access token from file if specified
    if args.token_file:
        try:
            with open(args.token_file, 'r') as f:
                args.access_token = f.read().strip()
        except Exception as e:
            print(f"Error reading token file: {e}")
            sys.exit(1)
    
    # Prompt for access token if not provided
    if not args.access_token:
        args.access_token = input("Enter your Amazon Quip access token (from https://quip-amazon.com/dev/token): ").strip()
        
    # Prompt for folder link if not provided
    if not args.folder_link:
        args.folder_link = input("Enter the Amazon Quip folder link to extract: ").strip()
    
    # Determine API URL based on the folder link domain if not specified
    if not args.api_url:
        domain = get_domain_from_link(args.folder_link)
        if domain:
            args.api_url = f"https://platform.{domain}"
        else:
            args.api_url = "https://platform.quip-amazon.com"  # Default for Amazon Quip
    
    # Initialize Quip client
    try:
        client = quip.QuipClient(access_token=args.access_token, base_url=args.api_url)
    except Exception as e:
        print(f"Failed to initialize Quip client: {e}")
        print("\nTips for getting a valid access token:")
        print("1. Go to https://quip-amazon.com/dev/token")
        print("2. Click 'Generate Personal Access Token'")
        print("3. Copy the token")
        print("4. Make sure you have the correct permissions")
        sys.exit(1)
    
    # Extract ID from link
    folder_id = extract_id_from_link(args.folder_link)
    if not folder_id:
        print("Failed to extract folder ID from the provided link.")
        print(f"The link provided was: {args.folder_link}")
        print("Please make sure you're using a valid Quip folder link.")
        sys.exit(1)
    
    # Start processing
    print(f"Starting extraction from folder ID: {folder_id}")
    print(f"Files will be saved to: {os.path.abspath(args.output)}")
    create_directory_if_not_exists(args.output)
    
    try:
        process_folder(client, folder_id, args.output)
        print("\nExtraction completed successfully!")
    except KeyboardInterrupt:
        print("\n\nExtraction was interrupted by user. Partial data may have been extracted.")
    except Exception as e:
        print(f"\nExtraction failed: {e}")


if __name__ == "__main__":
    main()