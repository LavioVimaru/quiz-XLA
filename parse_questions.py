import zipfile
import xml.etree.ElementTree as ET
import json
import re
import os
import sys

# Set default encoding to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

docx_path = "đề cương xla.docx"
output_json = "questions.json"
images_dir = "images"

namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'v': 'urn:schemas-microsoft-com:vml'
}

def extract_and_parse():
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)
        
    print("Opening ZIP file of đề cương xla.docx...")
    with zipfile.ZipFile(docx_path) as z:
        # 1. Extract and map images from relationships
        rels_content = z.read('word/_rels/document.xml.rels')
        rels_root = ET.fromstring(rels_content)
        ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        rels_map = {}
        for rel in rels_root.findall('.//r:Relationship', ns_rel):
            rels_map[rel.get('Id')] = rel.get('Target')
            
        # Extract images from ZIP
        print("Extracting images...")
        extracted_count = 0
        for name in z.namelist():
            if name.startswith('word/media/') and not name.endswith('/'):
                filename = name.split('/')[-1]
                if not filename:
                    continue
                target_path = os.path.join(images_dir, filename)
                # Read content and save
                img_data = z.read(name)
                with open(target_path, 'wb') as f:
                    f.write(img_data)
                extracted_count += 1
        print(f"Extracted {extracted_count} images into '{images_dir}/'.")
        
        # 2. Parse document XML
        xml_content = z.read('word/document.xml')
        root = ET.fromstring(xml_content)
        
        paragraphs = root.findall('.//w:p', namespaces)
        
        parsed_paragraphs = []
        for idx, p in enumerate(paragraphs):
            runs_info = []
            img_names = []
            
            # Check for drawings
            for d in p.findall('.//w:drawing', namespaces):
                blip = d.find('.//a:blip', namespaces)
                if blip is not None:
                    embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if embed in rels_map:
                        target = rels_map[embed]
                        img_names.append(target.split('/')[-1])
            
            # Check for VML images
            for im in p.findall('.//v:imagedata', namespaces):
                href = im.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}href')
                id_ = im.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                embed = id_ if id_ else href
                if embed in rels_map:
                    target = rels_map[embed]
                    img_names.append(target.split('/')[-1])
            
            for child in p:
                if child.tag.endswith('r'):
                    r = child
                    rPr = r.find('.//w:rPr', namespaces)
                    is_highlight = False
                    if rPr is not None:
                        if rPr.find('.//w:highlight', namespaces) is not None:
                            is_highlight = True
                    
                    for run_child in r:
                        if run_child.tag.endswith('t'):
                            if run_child.text:
                                runs_info.append({'text': run_child.text, 'highlight': is_highlight})
                        elif run_child.tag.endswith('tab'):
                            runs_info.append({'text': '\t', 'highlight': False})
                            
            p_text = "".join(r['text'] for r in runs_info)
            parsed_paragraphs.append({
                'index': idx,
                'runs': runs_info,
                'text': p_text,
                'images': img_names
            })
            
    # 3. Group paragraphs by questions using DP sequence check
    q_pattern = re.compile(r'^\s*(\d+)\s*\.(.*)', re.DOTALL)
    candidates = []
    for idx, p_data in enumerate(parsed_paragraphs):
        match = q_pattern.match(p_data['text'])
        if match:
            q_num = int(match.group(1))
            candidates.append({
                'p_idx': idx,
                'q_num': q_num,
                'text_parts': [match.group(2).strip()],
                'images': list(p_data['images'])
            })
            
    n = len(candidates)
    dp = [0] * n
    parent = [-1] * n
    
    # Initialize DP for candidates starting at q_num = 1
    for i in range(n):
        if candidates[i]['q_num'] == 1:
            dp[i] = 1
            
    for i in range(n):
        for j in range(i):
            if dp[j] > 0:
                diff = candidates[i]['q_num'] - candidates[j]['q_num']
                # Strictly increasing and difference <= 5
                if 0 < diff <= 5:
                    if dp[j] + 1 > dp[i]:
                        dp[i] = dp[j] + 1
                        parent[i] = j
                        
    # Find the candidate with the maximum DP value
    max_idx = -1
    max_val = 0
    for i in range(n):
        if dp[i] > max_val:
            max_val = dp[i]
            max_idx = i
            
    # Reconstruct correct sequence path
    correct_candidates = []
    curr = max_idx
    while curr != -1:
        correct_candidates.append(candidates[curr])
        curr = parent[curr]
    correct_candidates.reverse()
    
    # Group paragraphs between sequential correct questions
    questions = []
    for idx_c, c in enumerate(correct_candidates):
        start_p_idx = c['p_idx']
        end_p_idx = correct_candidates[idx_c + 1]['p_idx'] if idx_c + 1 < len(correct_candidates) else len(parsed_paragraphs)
        
        # Subparagraphs are the paragraphs between start_p_idx + 1 and end_p_idx - 1 (inclusive)
        subparagraphs = parsed_paragraphs[start_p_idx + 1:end_p_idx]
        
        questions.append({
            'question_num': c['q_num'],
            'question_text_parts': c['text_parts'],
            'images': c['images'],
            'subparagraphs': subparagraphs,
            'start_idx': start_p_idx
        })
        
    print(f"Grouped {len(questions)} questions.")
    
    # 4. Extract options and clean texts for each question
    cleaned_questions = []
    option_prefix_pattern = re.compile(r'^([A-H]\s*[\)\.\-\:]\s*|[1-8]\s*[\)\.\-\:]\s*(?!\d))', re.IGNORECASE)
    
    for q in questions:
        # Identify options from subparagraphs
        subps = [sp for sp in q['subparagraphs'] if sp['text'].strip() or sp['images']]
        
        raw_options = []
        question_additional_texts = []
        question_images = list(q['images'])
        
        # Check if this is a single-answer image question
        is_single_ans_image = False
        ans_sp_idx = -1
        for idx_sp, sp in enumerate(subps):
            if any(term in sp['text'].lower() for term in ['đáp án', 'dáp án', 'dap an']):
                if sp['images']:
                    is_single_ans_image = True
                    ans_sp_idx = idx_sp
                    break
                    
        if is_single_ans_image:
            for idx_sp in range(ans_sp_idx):
                text_part = subps[idx_sp]['text'].strip()
                if text_part:
                    question_additional_texts.append(text_part)
                for img in subps[idx_sp]['images']:
                    if img not in question_images:
                        question_images.append(img)
                        
            opt_images = [f"images/{img}" for img in subps[ans_sp_idx]['images']]
            cleaned_options = [{
                'text': 'Đáp án',
                'is_correct': True,
                'images': opt_images
            }]
            
            main_q_text = " ".join(q['question_text_parts']).strip()
            full_q_text = main_q_text
            if question_additional_texts:
                full_q_text += "\n" + "\n".join(question_additional_texts)
                
            q_img_paths = [f"images/{img}" for img in question_images]
            cleaned_questions.append({
                'question_num': q['question_num'],
                'question': f"Câu hỏi {q['question_num']}: {full_q_text}",
                'images': q_img_paths,
                'options': cleaned_options
            })
            continue
        
        for sp in subps:
            # Check if this paragraph is part of the question rather than options.
            # Usually, drawings inside non-option paragraphs or text before options.
            # But let's build the options first.
            runs = sp['runs']
            images = sp['images']
            
            curr_opt_text = []
            curr_opt_correct = False
            
            paragraph_options = []
            
            for r in runs:
                text = r['text']
                high = r['highlight']
                
                if '\t' in text:
                    parts = text.split('\t')
                    for idx, part in enumerate(parts):
                        if idx > 0:
                            opt_str = "".join(curr_opt_text).strip()
                            if opt_str:
                                paragraph_options.append({'text': opt_str, 'is_correct': curr_opt_correct, 'images': []})
                            curr_opt_text = []
                            curr_opt_correct = False
                        curr_opt_text.append(part)
                        if high:
                            curr_opt_correct = True
                else:
                    if re.search(r'\s{3,}', text):
                        parts = re.split(r'\s{3,}', text)
                        for idx, part in enumerate(parts):
                            if idx > 0:
                                opt_str = "".join(curr_opt_text).strip()
                                if opt_str:
                                    paragraph_options.append({'text': opt_str, 'is_correct': curr_opt_correct, 'images': []})
                                curr_opt_text = []
                                curr_opt_correct = False
                            curr_opt_text.append(part)
                            if high:
                                curr_opt_correct = True
                    else:
                        curr_opt_text.append(text)
                        if high:
                            curr_opt_correct = True
                            
            opt_str = "".join(curr_opt_text).strip()
            if opt_str:
                paragraph_options.append({'text': opt_str, 'is_correct': curr_opt_correct, 'images': list(images)})
            elif images:
                paragraph_options.append({'text': '', 'is_correct': curr_opt_correct, 'images': list(images)})
                
            # If this paragraph didn't split into multiple options, and has no highlight,
            # and we haven't seen any highlighted options yet, is it part of the question text?
            # Let's keep it as option for now, and clean it up later.
            raw_options.extend(paragraph_options)
            
        # Clean options:
        # Heuristic 1: If there are 5 options, and the first option has no highlight,
        # and it ends with ':' or is very short (like 'K-Mean:'), merge it into question text.
        if len(raw_options) == 5 and not raw_options[0]['is_correct'] and (raw_options[0]['text'].endswith(':') or len(raw_options[0]['text']) < 15):
            merged_part = raw_options[0]['text']
            question_additional_texts.append(merged_part)
            # If the merged option had images, add them to question images
            if raw_options[0]['images']:
                question_images.extend(raw_options[0]['images'])
            raw_options = raw_options[1:]
            
        # Heuristic 2: For options that have text, strip standard prefixes like 'A.', 'B)', '1.', etc.
        cleaned_options = []
        for opt in raw_options:
            cleaned_text = option_prefix_pattern.sub('', opt['text']).strip()
            if not cleaned_text:
                cleaned_text = opt['text'].strip()
            # If the text was just the prefix and nothing else, ignore or keep
            # Let's keep it unless it becomes empty (but if it's empty and has image, it's fine)
            opt_images = [f"images/{img}" for img in opt['images']]
            cleaned_options.append({
                'text': cleaned_text,
                'is_correct': opt['is_correct'],
                'images': opt_images
            })
            
        # Heuristic 3: If a question has only 1 option in total, mark it as correct
        if len(cleaned_options) == 1:
            cleaned_options[0]['is_correct'] = True
            
        # Special Hardcode: For Q250, force the 4th option '8' to be correct
        if q['question_num'] == 250 and not any(o['is_correct'] for o in cleaned_options):
            for o in cleaned_options:
                if o['text'] == '8':
                    o['is_correct'] = True
                    
        # Construct full question text
        main_q_text = " ".join(q['question_text_parts']).strip()
        # Find any other non-option text blocks to include in question
        # For example, paragraphs that contains question images or context (e.g. Q284: "Hỏi giá trị trung bình...")
        # If any paragraph in subparagraphs was not classified as an option (e.g., empty text but has images, or text that comes before options)
        # Actually, let's look at subparagraphs.
        # If there are subparagraphs that have text and no options were extracted (this doesn't happen because we treat all text as options, except Heuristic 1)
        # But wait! In Q284, the subparagraphs were:
        # P 1409: '' [Images: image34.png]
        # P 1410: 'Hỏi giá trị trung bình với cửa sổ lọc 3x3 bằng bao nhiêu'
        # P 1411: '48.7. [H]                               69' (options!)
        # Our previous simplified script treated P 1410 as an option ('Hỏi giá trị trung bình...'). But it's part of the question!
        # How do we know P 1410 is part of the question and not an option?
        # Let's look at P 1410: it does not end with any option format and is followed by actual options.
        # Let's define a robust check:
        # In a list of extracted raw_options:
        # Options are usually at the end of the subparagraph list.
        # If an option has no highlight, does not start with an option prefix,
        # is followed by other options, and is not a short text separator, it might be part of the question text.
        # But to be safe: let's see. If the question has more than 4 options (e.g. 6 options: "Hỏi giá trị...", "48.7", "69", "49", "48"),
        # we can identify which ones are actual options.
        # Actual options often match pattern like "starts with prefix" or are short numbers, or are highlighted.
        # Let's refine this:
        # If a question has > 4 options:
        # Let's identify the options. If the first options are long texts and the last 4 are actual options (often short or matching A,B,C,D),
        # we merge the non-options into the question text.
        # Let's print out what questions have > 4 options in our cleaned list to see if we need this.
        # We can implement a check: if options count > 4:
        # Let's see if we can identify non-options.
        # E.g. for Q284:
        # Question text: "Cho ảnh I có giá trị"
        # Subparagraphs:
        # - P 1409: (image34)
        # - P 1410: "Hỏi giá trị trung bình..." (raw option: "Hỏi giá trị trung bình...")
        # - P 1411: "48.7" (correct), "69"
        # - P 1412: "49", "48"
        # So raw options list has:
        # 0. "Hỏi giá trị trung bình..." (correct: False)
        # 1. "48.7" (correct: True)
        # 2. "69" (correct: False)
        # 3. "49" (correct: False)
        # 4. "48" (correct: False)
        # That's 5 options! Our Heuristic 1 checks if `raw_options[0]['text'].endswith(':')` or `len < 15`.
        # Here, "Hỏi giá trị trung bình với cửa sổ lọc 3x3 bằng bao nhiêu" has length 58 and doesn't end with ':'.
        # So it would not be merged by Heuristic 1! It would remain as an option.
        # To fix this: if we have 5 options, and the first option has no highlight, and it's a long sentence,
        # while the remaining 4 options are short values/numbers or typical option formats:
        # Yes! We can check if `len(raw_options) == 5` and the first option has no highlight,
        # and we can merge it if the other 4 options look like typical options (e.g. their average length is short, or they are numbers, or they start with prefixes).
        # Let's write a python check:
        # If `len(raw_options) == 5` and the first option is not correct:
        # Let's check average length of options 1-4. If it's small (e.g. < 30 characters), then the first option (which is likely a sub-question text) should be merged into the question text!
        # This is extremely smart and covers Q284 perfectly!
        
        if len(raw_options) == 5 and not raw_options[0]['is_correct']:
            avg_len_others = sum(len(o['text']) for o in raw_options[1:]) / 4
            if avg_len_others < 30:
                merged_part = raw_options[0]['text']
                question_additional_texts.append(merged_part)
                if raw_options[0]['images']:
                    question_images.extend(raw_options[0]['images'])
                raw_options = raw_options[1:]
                
        # Re-apply Heuristic 2 & 3 after potential merge
        cleaned_options = []
        for opt in raw_options:
            cleaned_text = option_prefix_pattern.sub('', opt['text']).strip()
            if not cleaned_text:
                cleaned_text = opt['text'].strip()
            opt_images = [f"images/{img}" for img in opt['images']]
            cleaned_options.append({
                'text': cleaned_text,
                'is_correct': opt['is_correct'],
                'images': opt_images
            })
        if len(cleaned_options) == 1:
            cleaned_options[0]['is_correct'] = True
            
        # Re-apply Q250 hardcode
        if q['question_num'] == 250 and not any(o['is_correct'] for o in cleaned_options):
            for o in cleaned_options:
                if o['text'] == '8':
                    o['is_correct'] = True
                    
        # Append additional question texts
        full_q_text = main_q_text
        if question_additional_texts:
            full_q_text += "\n" + "\n".join(question_additional_texts)
            
        # Associate question images
        # Collect images from subparagraphs that are not options, or question_images
        # Let's collect all images that are in subparagraphs but not in the options
        q_img_paths = []
        for img in question_images:
            q_img_paths.append(f"images/{img}")
        # Add other subparagraph images that are not associated with any option
        for sp in subps:
            for img in sp['images']:
                img_path = f"images/{img}"
                # If this image is not in any of the cleaned options, it belongs to the question
                is_opt_img = False
                for opt in cleaned_options:
                    if img_path in opt['images']:
                        is_opt_img = True
                        break
                if not is_opt_img and img_path not in q_img_paths:
                    q_img_paths.append(img_path)
                    
        cleaned_questions.append({
            'question_num': q['question_num'],
            'question': f"Câu hỏi {q['question_num']}: {full_q_text}",
            'images': q_img_paths,
            'options': cleaned_options
        })
        
    # Write questions list to JSON file
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(cleaned_questions, f, ensure_ascii=False, indent=2)
        
    print(f"Successfully generated {output_json} with {len(cleaned_questions)} questions.")

if __name__ == '__main__':
    extract_and_parse()
