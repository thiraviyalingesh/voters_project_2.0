#!/usr/bin/env python3
"""
Voter Analytics - Headless Batch Processor v8.0 (Ubuntu VM Optimized)
Fixed: Deadlock prevention on high-core Linux systems using Spawn context.
"""

import os
# --- CRITICAL: MUST BE BEFORE ANY OTHER IMPORTS ---
# Prevents Tesseract/OpenMP thread explosion (8 cores * 8 threads = 64 threads = HANG)
os.environ["OMP_THREAD_LIMIT"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["OPENBLAS_NUM_THREADS"] = "1"

import re
import sys
import argparse
from pathlib import Path
import time
import json
import requests
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed

# Import image processing libraries
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import pytesseract

# Configuration
NTFY_TOPIC = None
# Leave 1 core free for system/IO to prevent VM hanging
NUM_WORKERS = max(1, (os.cpu_count() or 8) - 1)

# --- WORKER FUNCTIONS (Must be at top level for Spawn) ---

def ocr_single_card(args):
    """Worker function to OCR a single voter card image."""
    jpg_path, global_idx, pdf_name = args
    try:
        img = Image.open(jpg_path)
        # Use config to further ensure single-threaded Tesseract
        custom_config = r'--oem 3 --psm 6 -c omp_thread_limit=1'
        text = pytesseract.image_to_string(img, lang='tam', config=custom_config)
        data = parse_voter_card(text)
        stem = Path(jpg_path).stem
        return global_idx, stem, data, pdf_name
    except Exception as e:
        return global_idx, str(global_idx), None, pdf_name

def enhanced_ocr_age_gender(args):
    """Worker function for enhanced OCR focused on Age and Gender only."""
    jpg_path, global_idx = args
    try:
        img = Image.open(str(jpg_path))
        width, height = img.size
        bottom_crop = img.crop((0, int(height * 0.65), width, height))
        
        # Reduced approaches to prevent CPU thrashing
        approaches = [
            ('original', lambda i: i),
            ('contrast', lambda i: ImageEnhance.Contrast(i).enhance(2.0)),
            ('scale_2x', lambda i: i.resize((i.size[0] * 2, i.size[1] * 2), Image.LANCZOS)),
        ]

        result = {'age': '', 'gender': ''}
        custom_config = r'--oem 3 --psm 6 -c omp_thread_limit=1'

        for name, transform in approaches:
            try:
                processed_img = transform(bottom_crop)
                text = pytesseract.image_to_string(processed_img, lang='tam+eng', config=custom_config)
                
                if not result['age']:
                    age_match = re.search(r'வயது\s*:\s*(\d+)', text)
                    if age_match: result['age'] = age_match.group(1)

                if not result['gender']:
                    if 'பாலினம்' in text:
                        if 'ஆண்' in text: result['gender'] = 'Male'
                        elif 'பெண்' in text: result['gender'] = 'Female'
                
                if result['age'] and result['gender']: break
            except: continue
        return global_idx, result
    except Exception:
        return global_idx, None

# --- UTILITY FUNCTIONS ---

def parse_voter_card(text):
    data = {'serial_no': '', 'voter_id': '', 'name': '', 'relation_name': '', 
            'relation_type': '', 'house_no': '', 'age': '', 'gender': ''}
    
    voter_id_patterns = [r'\b([A-Z]{2,3}\d{6,10})\b', r'\b([A-Z0-9]{2,3}\d{6,10})\b']
    for pattern in voter_id_patterns:
        match = re.search(pattern, text)
        if match and len(match.group(1)) >= 9:
            data['voter_id'] = match.group(1)
            break

    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        
        if not data['serial_no']:
            serial_match = re.match(r'^(\d{1,4})\b', line)
            if serial_match: data['serial_no'] = serial_match.group(1)

        if 'பெயர்' in line and ':' in line:
            name_part = line.split(':', 1)[-1].strip()
            if 'தந்தை' not in line and 'கணவர்' not in line:
                if not data['name']: data['name'] = name_part
        
        if 'தந்தை' in line: 
            data['relation_name'] = line.split(':', 1)[-1].strip()
            data['relation_type'] = 'Father'
        elif 'கணவர்' in line:
            data['relation_name'] = line.split(':', 1)[-1].strip()
            data['relation_type'] = 'Husband'
        
        if 'வயது' in line:
            age_m = re.search(r'(\d+)', line)
            if age_m: data['age'] = age_m.group(1)
        
        if 'பாலினம்' in line:
            if 'ஆண்' in line: data['gender'] = 'Male'
            elif 'பெண்' in line: data['gender'] = 'Female'

    return data

def extract_cards_from_pdf_sequential(pdf_path, temp_base_dir, pdf_index):
    try:
        pdf_path = Path(pdf_path)
        pdf_name = pdf_path.stem
        output_path = Path(temp_base_dir) / pdf_name
        output_path.mkdir(parents=True, exist_ok=True)

        doc = fitz.open(str(pdf_path))
        card_count = 0
        for page_num in range(3, len(doc) - 1):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            page_img = Image.open(io.BytesIO(pix.tobytes("png")))
            
            w, h = page_img.size
            c_w, r_h = w // 3, (h - int(h*0.06)) // 10
            
            for r in range(10):
                for c in range(3):
                    x1, y1 = c * c_w, int(h*0.035) + r * r_h
                    card_img = page_img.crop((x1+2, y1+2, x1+c_w-2, y1+r_h-2))
                    
                    # Skip empty white cards
                    if sum(card_img.getpixel((10,10))) / 3 > 250: continue
                    
                    card_count += 1
                    card_img.save(output_path / f"{card_count}.png")
        doc.close()
        return pdf_index, pdf_name, card_count, str(output_path)
    except Exception as e:
        print(f"Error: {e}")
        return pdf_index, Path(pdf_path).stem, 0, None

# --- CORE LOGIC ---

class CheckpointManager:
    def __init__(self, checkpoint_path):
        self.checkpoint_path = Path(checkpoint_path)
        self.data = {'phase': 0, 'extracted_pdfs': {}, 'ocr_results': {}, 'all_cards': []}
    def load(self):
        if self.checkpoint_path.exists():
            with open(self.checkpoint_path, 'r') as f: self.data = json.load(f)
            return True
        return False
    def save(self):
        with open(self.checkpoint_path, 'w') as f: json.dump(self.data, f)
    def delete(self):
        if self.checkpoint_path.exists(): self.checkpoint_path.unlink()

def process_constituency(folder_path, ntfy_topic=None):
    folder_path = Path(folder_path)
    constituency_name = folder_path.name
    pdf_files = sorted(folder_path.glob("*.pdf"))
    
    checkpoint = CheckpointManager(folder_path.parent / f".{constituency_name}_cp.json")
    image_dir = folder_path.parent / f".{constituency_name}_tmp"
    image_dir.mkdir(parents=True, exist_ok=True)
    
    checkpoint.load()
    
    # Phase 1: PDF Extraction
    if checkpoint.data['phase'] < 1:
        print(f"Phase 1: Extracting {len(pdf_files)} PDFs...")
        all_cards = []
        for idx, pdf in enumerate(pdf_files):
            _, name, count, out = extract_cards_from_pdf_sequential(pdf, image_dir, idx)
            if out:
                for i in range(1, count + 1):
                    all_cards.append((str(Path(out)/f"{i}.png"), i, name))
        checkpoint.data['all_cards'] = all_cards
        checkpoint.data['phase'] = 1
        checkpoint.save()

    # Phase 2: OCR with Spawn Context (The Fix for Hanging)
    if checkpoint.data['phase'] < 2:
        print(f"Phase 2: OCR with {NUM_WORKERS} workers...")
        ctx = multiprocessing.get_context('spawn')
        cards = checkpoint.data['all_cards']
        ocr_results = {}
        
        with ProcessPoolExecutor(max_workers=NUM_WORKERS, mp_context=ctx) as executor:
            futures = {executor.submit(ocr_single_card, c): i for i, c in enumerate(cards)}
            for f in as_completed(futures):
                g_idx, s_no, data, pdf_n = f.result()
                ocr_results[g_idx] = (s_no, data, pdf_n)
                if len(ocr_results) % 100 == 0: print(f"  Processed {len(ocr_results)} cards...")
        
        checkpoint.data['ocr_results'] = ocr_results
        checkpoint.data['phase'] = 2
        checkpoint.save()

    # Phase 4: Excel
    print("Phase 4: Generating Excel...")
    wb = Workbook()
    ws = wb.active
    ws.append(['S.No', 'Part', 'VoterID', 'Name', 'Age', 'Gender', 'PDF'])
    
    res = checkpoint.data['ocr_results']
    for k in sorted(res.keys(), key=int):
        s, d, p = res[k]
        if d:
            ws.append([s, p, d.get('voter_id'), d.get('name'), d.get('age'), d.get('gender'), p])
    
    out_file = folder_path.parent / f"{constituency_name}_results.xlsx"
    wb.save(out_file)
    print(f"Done! Saved to {out_file}")
    checkpoint.delete()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder')
    args = parser.parse_args()
    process_constituency(args.folder)

if __name__ == "__main__":
    main()