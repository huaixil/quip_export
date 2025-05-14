# Quip Export Tool

A command-line tool to export Quip documents and folders with their hierarchical structure, preserving attachments and images.

## Features

- Export entire folder hierarchies from Quip
- Download documents as DOCX files (with HTML fallback)
- Preserve folder structure
- Extract all attachments
- Maintain document titles and organization

## Requirements

- Python 3.6+
- Required packages:
  - requests
  - quip-api (included in the repository)

## Installation

1. Clone this repository or download the files
2. Install required packages:
   ```
   pip install requests
   ```

## Usage

```
python quip_export.py [options]
```

### Options

- `--token`, `-t`: Your Quip access token
- `--folder`, `-f`: Link to the Quip folder to export
- `--output`, `-o`: Output directory (default: ./quip_export)
- `--token-file`: Path to file containing your Quip access token
- `--api-url`: Custom Quip API URL (default: https://platform.quip-amazon.com)

### Examples

Export a folder using an access token:
```
python quip_export.py -t YOUR_ACCESS_TOKEN -f https://quip-amazon.com/folder/YOUR_FOLDER_ID
```

Export a folder using a token stored in a file:
```
python quip_export.py --token-file token.txt -f https://quip-amazon.com/folder/YOUR_FOLDER_ID -o ./exported_docs
```

## Getting a Quip Access Token

1. Go to https://quip-amazon.com/dev/token
2. Click 'Generate Personal Access Token'
3. Copy the token
4. Use it with the `--token` option or save it to a file for use with `--token-file`

## Output Structure

The tool creates a directory structure that mirrors your Quip folders:

```
output_directory/
├── Folder1/
│   ├── Document1.docx
│   ├── Document2.docx
│   └── Document2_attachments/
│       └── attachment.pdf
└── Folder2/
    ├── Document3.docx
    └── Subfolder/
        └── Document4.docx
```