# EML to TXT Converter

A powerful Python utility for converting email files (.eml) to human-readable text files and extracting email attachments.

## Features

- **Email Conversion**: Converts .eml files to readable .txt files
- **Attachment Extraction**: Extracts all email attachments to a dedicated folder
- **International Support**: Properly handles international character sets and encoded headers
- **Content Handling**: Extracts plain text content and notes HTML content
- **Recursive Processing**: Option to process nested folder structures
- **Preserves Organization**: Maintains folder hierarchies when processing recursively

## Installation

No external dependencies required! The script uses only Python standard library modules.

```bash
# Clone the repository
git clone https://github.com/praschak/eml-to-txt-converter.git
cd eml-to-txt-converter

# Make the script executable
chmod +x eml_to_txt.py
```

## Usage

### Basic Usage

```bash
# Convert all EML files in the current directory
python eml_to_txt.py .

# Convert and extract attachments
python eml_to_txt.py . --extract
```

### Advanced Options

```bash
# Full usage information
python eml_to_txt.py --help

# Specify output directory
python eml_to_txt.py ./emails -o ./converted

# Extract attachments to custom folder
python eml_to_txt.py ./emails -e -a ./attachments

# Process directories recursively
python eml_to_txt.py ./email_archive -r -o ./readable_emails -e
```

### Command Line Arguments

| Argument | Short | Description |
|----------|-------|-------------|
| `--output` | `-o` | Output folder for text files (default: same as input) |
| `--attachments` | `-a` | Folder for extracted attachments (default: 'attachments' subfolder) |
| `--recursive` | `-r` | Process subfolders recursively |
| `--extract` | `-e` | Extract attachments from emails |

## Output Format

### Text Files

The converted text files maintain the structure of the original emails with clear sections:

- Email headers (From, To, Subject, Date, etc.)
- Body content in plain text
- Information about attachments, including where they were saved

Example:
```
================================================================================
EMAIL: example.eml
================================================================================

From: sender@example.com
To: recipient@example.com
Subject: Meeting Notes
Date: Sat, 8 Mar 2025 10:00:00 -0500

--------------------------------------------------------------------------------

BODY:

Hi Team,

Attached are the notes from yesterday's meeting.

Please review and let me know if you have any questions.

Best regards,
John

--------------------------------------------------------------------------------
ATTACHMENTS:
[ATTACHMENT: meeting_notes.pdf (application/pdf, ~256.5 KB)] - Saved as: example_meeting_notes.pdf
```

### Attachments

Attachments are saved with sanitized filenames, prefixed with the email name to maintain the connection to their source email.

## Use Cases

- **Email Archiving**: Convert email archives to searchable, readable text
- **Data Migration**: Extract email content and attachments from legacy systems
- **Forensic Analysis**: Examine email content in a standardized format
- **Backup Solutions**: Create human-readable backups of important emails

## Requirements

- Python 3.6 or higher
