#!/usr/bin/env python3
"""
EML to TXT Converter - Converts all .eml files in a directory to human-readable .txt files
and extracts all attachments
"""

import os
import sys
import email
import base64
import quopri
from email.header import decode_header
import argparse
from pathlib import Path
import mimetypes


def decode_content(content, encoding):
    """Decode content based on Content-Transfer-Encoding"""
    if encoding == 'base64':
        try:
            return base64.b64decode(content).decode('utf-8', errors='replace')
        except:
            return "[BASE64 ENCODED BINARY CONTENT - NOT DISPLAYED]"
    elif encoding == 'quoted-printable':
        return quopri.decodestring(content).decode('utf-8', errors='replace')
    elif encoding == '7bit' or encoding == '8bit' or encoding is None:
        return content
    else:
        return f"[UNKNOWN ENCODING: {encoding}]"


def decode_header_value(header_value):
    """Decode email header values that might be encoded"""
    if header_value is None:
        return ""
    
    decoded_parts = []
    for part, encoding in decode_header(header_value):
        if isinstance(part, bytes):
            if encoding:
                try:
                    decoded_parts.append(part.decode(encoding, errors='replace'))
                except LookupError:
                    decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(part.decode('utf-8', errors='replace'))
        else:
            decoded_parts.append(part)
    
    return ' '.join(decoded_parts)


def sanitize_filename(filename):
    """Sanitize filename to avoid issues with invalid characters"""
    # Replace characters that might be problematic in filenames
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Limit filename length
    if len(filename) > 200:
        name, ext = os.path.splitext(filename)
        filename = name[:195] + ext
        
    return filename


def extract_attachment(part, attachments_dir, email_name):
    """Extract attachment from email part and save to disk"""
    filename = part.get_filename()
    if not filename:
        # Generate a filename if none is provided
        content_type = part.get_content_type()
        ext = mimetypes.guess_extension(content_type) or '.bin'
        filename = f"unnamed_attachment_{content_type.replace('/', '_')}{ext}"
    
    # Decode and sanitize the filename
    filename = decode_header_value(filename)
    filename = sanitize_filename(filename)
    
    # Create a unique filename with email name prefix to avoid conflicts
    email_prefix = email_name.split('.')[0]
    unique_filename = f"{email_prefix}_{filename}"
    file_path = Path(attachments_dir) / unique_filename
    
    # Extract and save the attachment
    payload = part.get_payload(decode=True)
    if payload:
        with open(file_path, 'wb') as f:
            f.write(payload)
        
        return {
            'filename': unique_filename,
            'original_filename': filename,
            'size': len(payload),
            'content_type': part.get_content_type(),
            'path': file_path
        }
    return None


def get_attachment_info(part, attachments_dir=None, email_name=None, extract=False):
    """Extract attachment information and optionally save attachment to disk"""
    if part.get_filename() or part.get_content_disposition() in ['attachment', 'inline']:
        content_type = part.get_content_type()
        
        # Skip if this is part of the email structure, not a real attachment
        if content_type in ['multipart/alternative', 'multipart/related', 'multipart/mixed']:
            return None
            
        filename = part.get_filename()
        if not filename:
            ext = mimetypes.guess_extension(content_type) or '.bin'
            filename = f"unnamed_attachment{ext}"
        else:
            filename = decode_header_value(filename)
        
        size = len(part.get_payload(decode=False))
        
        # Extract attachment if requested
        extracted_info = None
        if extract and attachments_dir and email_name:
            extracted_info = extract_attachment(part, attachments_dir, email_name)
        
        info = f"[ATTACHMENT: {filename} ({content_type}, ~{size/1024:.1f} KB)]"
        if extracted_info:
            info += f" - Saved as: {extracted_info['filename']}"
        
        return {
            'info_text': info,
            'extracted': extracted_info
        }
    return None


def process_eml_file(eml_path, output_path=None, attachments_dir=None, extract_attachments=False):
    """Process a single EML file and convert it to human-readable text"""
    if output_path is None:
        output_path = eml_path.with_suffix('.txt')
    
    # Create attachments directory if needed
    if extract_attachments and attachments_dir:
        Path(attachments_dir).mkdir(parents=True, exist_ok=True)
    
    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f)
    
    # Start with headers
    output = []
    output.append("=" * 80)
    output.append(f"EMAIL: {eml_path.name}")
    output.append("=" * 80)
    output.append("")
    
    # Process key headers with proper decoding
    important_headers = ['From', 'To', 'Cc', 'Bcc', 'Subject', 'Date']
    for header in important_headers:
        if header in msg:
            decoded_value = decode_header_value(msg[header])
            output.append(f"{header}: {decoded_value}")
    
    output.append("")
    output.append("-" * 80)
    output.append("")
    
    # Process body and attachments
    body_text = []
    attachment_info = []
    extracted_attachments = []
    
    # Function to process message parts recursively
    def extract_parts(message_part):
        content_type = message_part.get_content_type()
        content_disposition = str(message_part.get('Content-Disposition', ''))
        
        # Handle attachments
        is_attachment = 'attachment' in content_disposition or 'inline' in content_disposition
        if is_attachment or (message_part.get_filename() and content_type not in ['multipart/mixed', 'multipart/alternative']):
            attachment = get_attachment_info(
                message_part, 
                attachments_dir=attachments_dir, 
                email_name=eml_path.name,
                extract=extract_attachments
            )
            if attachment:
                attachment_info.append(attachment['info_text'])
                if attachment.get('extracted'):
                    extracted_attachments.append(attachment['extracted'])
            return
        
        # Handle message body parts
        if content_type == 'text/plain':
            content = message_part.get_payload(decode=True)
            charset = message_part.get_content_charset() or 'utf-8'
            encoding = message_part.get('Content-Transfer-Encoding')
            
            try:
                decoded_text = content.decode(charset, errors='replace')
                body_text.append(decoded_text)
            except:
                decoded_text = decode_content(content, encoding)
                body_text.append(decoded_text)
                
        elif content_type == 'text/html':
            # Just note that there was HTML content
            body_text.append("\n[HTML CONTENT AVAILABLE BUT NOT DISPLAYED]\n")
            
        # Process multipart messages recursively
        elif message_part.is_multipart():
            for part in message_part.get_payload():
                extract_parts(part)
    
    # Start extraction
    if msg.is_multipart():
        for part in msg.get_payload():
            extract_parts(part)
    else:
        extract_parts(msg)
    
    # Add body to output
    if body_text:
        output.append("BODY:")
        output.append("")
        output.extend(body_text)
    
    # Add attachment information
    if attachment_info:
        output.append("")
        output.append("-" * 80)
        output.append("ATTACHMENTS:")
        output.extend(attachment_info)
    
    # Write to output file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output))
    
    return {
        'output_path': output_path,
        'extracted_attachments': extracted_attachments
    }


def convert_folder(folder_path, output_folder=None, attachments_folder=None, recursive=False, extract_attachments=False):
    """Convert all .eml files in a folder to .txt files"""
    folder_path = Path(folder_path).resolve()
    
    if not folder_path.exists() or not folder_path.is_dir():
        print(f"Error: {folder_path} is not a valid directory")
        return False
    
    # Set up output folder
    if output_folder:
        output_folder = Path(output_folder).resolve()
        output_folder.mkdir(parents=True, exist_ok=True)
    
    # Set up attachments folder if extracting
    if extract_attachments:
        if attachments_folder:
            attachments_folder = Path(attachments_folder).resolve()
        else:
            # Default to 'attachments' subfolder in output folder
            if output_folder:
                attachments_folder = output_folder / 'attachments'
            else:
                attachments_folder = folder_path / 'attachments'
        
        attachments_folder.mkdir(parents=True, exist_ok=True)
    else:
        attachments_folder = None
    
    # Find all .eml files
    pattern = '**/*.eml' if recursive else '*.eml'
    eml_files = list(folder_path.glob(pattern))
    
    if not eml_files:
        print(f"No .eml files found in {folder_path}")
        return False
    
    print(f"Found {len(eml_files)} .eml files to convert")
    
    # Process each file
    processed_files = []
    total_attachments = 0
    
    for eml_file in eml_files:
        try:
            # Determine output path
            if output_folder:
                rel_path = eml_file.relative_to(folder_path)
                # Handle subdirectories if recursive
                if recursive and rel_path.parent != Path('.'):
                    target_dir = output_folder / rel_path.parent
                    target_dir.mkdir(parents=True, exist_ok=True)
                    
                    # Create subdirectory in attachments folder too if needed
                    if extract_attachments and attachments_folder:
                        attach_target_dir = attachments_folder / rel_path.parent
                        attach_target_dir.mkdir(parents=True, exist_ok=True)
                        current_attachments_dir = attach_target_dir
                    else:
                        current_attachments_dir = attachments_folder
                else:
                    target_dir = output_folder
                    current_attachments_dir = attachments_folder
                
                output_path = target_dir / f"{eml_file.stem}.txt"
            else:
                output_path = eml_file.with_suffix('.txt')
                current_attachments_dir = attachments_folder
            
            result = process_eml_file(
                eml_file, 
                output_path, 
                attachments_dir=current_attachments_dir,
                extract_attachments=extract_attachments
            )
            
            processed_files.append(result['output_path'])
            num_attachments = len(result['extracted_attachments'])
            total_attachments += num_attachments
            
            attachment_msg = f" ({num_attachments} attachments extracted)" if extract_attachments and num_attachments > 0 else ""
            print(f"Converted: {eml_file.name} → {output_path.name}{attachment_msg}")
            
        except Exception as e:
            print(f"Error processing {eml_file}: {str(e)}")
    
    if extract_attachments:
        print(f"\nTotal attachments extracted: {total_attachments}")
    
    print(f"\nSuccessfully converted {len(processed_files)} out of {len(eml_files)} files")
    return True


def main():
    parser = argparse.ArgumentParser(description='Convert EML files to human-readable TXT files')
    parser.add_argument('folder', help='Folder containing .eml files')
    parser.add_argument('-o', '--output', help='Output folder for TXT files (default: same as input)')
    parser.add_argument('-a', '--attachments', help='Folder for extracted attachments (default: attachments subfolder)')
    parser.add_argument('-r', '--recursive', action='store_true', help='Process subfolders recursively')
    parser.add_argument('-e', '--extract', action='store_true', help='Extract attachments from emails')
    args = parser.parse_args()
    
    print("\nEML to TXT Converter")
    print("===================\n")
    
    if convert_folder(
        args.folder, 
        args.output, 
        args.attachments, 
        args.recursive, 
        args.extract
    ):
        print("\nConversion completed successfully!")
        return 0
    else:
        print("\nConversion failed or no files were processed.")
        return 1


if __name__ == "__main__":
    sys.exit(main())