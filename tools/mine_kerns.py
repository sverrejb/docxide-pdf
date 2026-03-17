"""Mine inter-glyph adjustments from a reference PDF's TJ arrays.

Usage: uv run tools/mine_kerns.py <reference.pdf> [--verbose]

Outputs a JSON kern database to stdout, stats to stderr.
Uses mutool to dump decompressed content streams.
"""
import sys
import json
import re
import subprocess


def main():
    pdf_path = sys.argv[1]
    verbose = "--verbose" in sys.argv

    # Get page count and object structure via mutool
    xref = run_mutool(pdf_path, "xref")
    pages = find_page_objects(pdf_path, xref)

    all_adjustments = {}  # (left_char, right_char) -> [adj_values]

    for page_id in pages:
        page_data = run_mutool(pdf_path, str(page_id))
        font_map = extract_font_map(pdf_path, page_data)
        content_id = extract_content_id(page_data)
        if not content_id:
            continue

        content = run_mutool(pdf_path, str(content_id))
        # Strip the object header (everything before "stream")
        stream_start = content.find("stream\n")
        if stream_start >= 0:
            content = content[stream_start + 7:]
        endstream = content.find("endstream")
        if endstream >= 0:
            content = content[:endstream]

        parse_content(content, font_map, all_adjustments, verbose)

    # Compute mean adjustment per pair
    kern_db = {}
    zero_count = 0
    nonzero_count = 0
    for (left, right), values in sorted(all_adjustments.items()):
        mean_val = sum(values) / len(values)
        if abs(mean_val) < 0.1:
            zero_count += 1
        else:
            kern_db[f"{left},{right}"] = round(mean_val, 2)
            nonzero_count += 1

    json.dump(kern_db, sys.stdout, indent=2, ensure_ascii=False)
    print(file=sys.stdout)
    print(f"{nonzero_count} non-zero pairs, {zero_count} zero pairs, "
          f"{len(all_adjustments)} total measured", file=sys.stderr)


def run_mutool(pdf_path, arg):
    result = subprocess.run(
        ["mutool", "show", pdf_path, arg],
        capture_output=True, text=True
    )
    return result.stdout


def find_page_objects(pdf_path, xref_text):
    """Find all page object IDs."""
    pages = []
    # Parse xref to get object count
    lines = xref_text.strip().split("\n")
    obj_count = 0
    for line in lines:
        parts = line.split()
        if len(parts) >= 2 and parts[0].isdigit():
            obj_count = max(obj_count, int(parts[0]))

    for obj_id in range(1, obj_count + 1):
        data = run_mutool(pdf_path, str(obj_id))
        if "/Type /Page" in data and "/Type /Pages" not in data:
            pages.append(obj_id)
    return pages


def extract_font_map(pdf_path, page_data):
    """Build font_key -> to_unicode mapping from page resources."""
    font_map = {}

    # Find font references: /TT2 7 0 R
    for m in re.finditer(r'/(\w+)\s+(\d+)\s+\d+\s+R', page_data):
        font_key = m.group(1)
        font_id = int(m.group(2))
        font_data = run_mutool(pdf_path, str(font_id))

        if "/Type /Font" not in font_data:
            continue

        # Get ToUnicode reference
        tu_match = re.search(r'/ToUnicode\s+(\d+)\s+\d+\s+R', font_data)
        if not tu_match:
            continue

        tu_id = int(tu_match.group(1))
        tu_data = run_mutool(pdf_path, str(tu_id))

        to_unicode = parse_cmap(tu_data)
        if to_unicode:
            font_map[font_key] = to_unicode

    return font_map


def extract_content_id(page_data):
    m = re.search(r'/Contents\s+(\d+)\s+\d+\s+R', page_data)
    if m:
        return int(m.group(1))
    # Array of content streams — take first
    m = re.search(r'/Contents\s*\[\s*(\d+)', page_data)
    if m:
        return int(m.group(1))
    return None


def parse_cmap(text):
    """Parse ToUnicode CMap from mutool dump."""
    mapping = {}
    lines = text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if 'beginbfchar' in line:
            i += 1
            while i < len(lines):
                line = lines[i].strip()
                if 'endbfchar' in line:
                    break
                parts = line.split()
                if len(parts) >= 2:
                    src = int(parts[0].strip('<>'), 16)
                    dst_hex = parts[1].strip('<>')
                    dst = decode_unicode_hex(dst_hex)
                    if dst:
                        mapping[src] = dst
                i += 1
        elif 'beginbfrange' in line:
            i += 1
            while i < len(lines):
                line = lines[i].strip()
                if 'endbfrange' in line:
                    break
                parts = line.split()
                if len(parts) >= 3:
                    start = int(parts[0].strip('<>'), 16)
                    end = int(parts[1].strip('<>'), 16)
                    dst_start = int(parts[2].strip('<>'), 16)
                    for code in range(start, end + 1):
                        ch = chr(dst_start + code - start)
                        mapping[code] = ch
                i += 1
        i += 1
    return mapping


def decode_unicode_hex(s):
    if len(s) <= 4:
        return chr(int(s, 16))
    result = ""
    for j in range(0, len(s), 4):
        if j + 4 <= len(s):
            result += chr(int(s[j:j+4], 16))
    return result


def parse_content(content, font_map, all_adjustments, verbose):
    """Parse content stream and extract TJ adjustments."""
    current_font_key = None
    current_to_unicode = None

    # Find all TJ arrays with their preceding Tf
    # The content is a single big string — parse operators
    pos = 0
    while pos < len(content):
        # Skip whitespace
        while pos < len(content) and content[pos] in ' \t\n\r':
            pos += 1
        if pos >= len(content):
            break

        # Look for key operators
        # Tf: /FontName size Tf
        tf_match = re.match(r'/(\w+)\s+[\d.]+\s+Tf', content[pos:])
        if tf_match:
            current_font_key = tf_match.group(1)
            current_to_unicode = font_map.get(current_font_key)
            pos += tf_match.end()
            continue

        # TJ: [...] TJ
        if content[pos] == '[':
            # Find matching ]
            bracket_depth = 1
            end = pos + 1
            in_string = False
            while end < len(content) and bracket_depth > 0:
                ch = content[end]
                if in_string:
                    if ch == '\\':
                        end += 1  # skip escaped char
                    elif ch == ')':
                        in_string = False
                elif ch == '(':
                    in_string = True
                elif ch == '[':
                    bracket_depth += 1
                elif ch == ']':
                    bracket_depth -= 1
                end += 1

            # Check if followed by TJ
            rest = content[end:end+10].lstrip()
            if rest.startswith('TJ'):
                array_content = content[pos+1:end-1]
                if current_to_unicode:
                    process_tj_array(array_content, current_to_unicode, all_adjustments, verbose)
                pos = end + 2
                continue

        # Skip to next whitespace/operator
        pos += 1
        while pos < len(content) and content[pos] not in ' \t\n\r/[':
            pos += 1


def process_tj_array(array_str, to_unicode, all_adjustments, verbose):
    """Process a TJ array string like '(!) -0.3 (") -0.1 ...'"""
    items = []  # list of ('char', unicode_char) or ('adjust', value)

    pos = 0
    while pos < len(array_str):
        ch = array_str[pos]
        if ch == '(':
            # Literal string — parse bytes
            end = pos + 1
            bytes_list = []
            while end < len(array_str) and array_str[end] != ')':
                if array_str[end] == '\\':
                    end += 1
                    if end < len(array_str):
                        esc = array_str[end]
                        if esc == 'n':
                            bytes_list.append(0x0a)
                        elif esc == 'r':
                            bytes_list.append(0x0d)
                        elif esc == 't':
                            bytes_list.append(0x09)
                        elif esc == '(':
                            bytes_list.append(0x28)
                        elif esc == ')':
                            bytes_list.append(0x29)
                        elif esc == '\\':
                            bytes_list.append(0x5c)
                        elif esc.isdigit():
                            # Octal
                            octal = esc
                            for _ in range(2):
                                if end + 1 < len(array_str) and array_str[end+1].isdigit():
                                    end += 1
                                    octal += array_str[end]
                            bytes_list.append(int(octal, 8))
                        else:
                            bytes_list.append(ord(esc))
                else:
                    bytes_list.append(ord(array_str[end]))
                end += 1
            # Map bytes to unicode via ToUnicode
            for b in bytes_list:
                unicode_ch = to_unicode.get(b)
                if unicode_ch:
                    items.append(('char', unicode_ch))
            pos = end + 1
        elif ch == '<':
            # Hex string
            end = array_str.find('>', pos)
            if end < 0:
                break
            hex_str = array_str[pos+1:end].replace(' ', '')
            for j in range(0, len(hex_str), 4):
                if j + 4 <= len(hex_str):
                    gid = int(hex_str[j:j+4], 16)
                    unicode_ch = to_unicode.get(gid)
                    if unicode_ch:
                        items.append(('char', unicode_ch))
            pos = end + 1
        elif ch in '-0123456789.':
            # Numeric adjustment
            end = pos + 1
            while end < len(array_str) and array_str[end] in '-0123456789.eE':
                end += 1
            try:
                val = float(array_str[pos:end])
                items.append(('adjust', val))
            except ValueError:
                pass
            pos = end
        else:
            pos += 1

    # Walk items and record adjustments
    prev_char = None
    i = 0
    while i < len(items):
        item = items[i]
        if item[0] == 'char':
            ch = item[1]
            if prev_char is not None:
                # No explicit adjustment between prev and this = implicit 0
                key = (prev_char, ch)
                all_adjustments.setdefault(key, []).append(0.0)
            prev_char = ch
        elif item[0] == 'adjust':
            adj = item[1]
            # Find the next char
            next_char = None
            for j in range(i + 1, len(items)):
                if items[j][0] == 'char':
                    next_char = items[j][1]
                    break
            if prev_char is not None and next_char is not None:
                key = (prev_char, next_char)
                # Remove the implicit 0 we just added and replace with actual adjustment
                vals = all_adjustments.setdefault(key, [])
                if vals and vals[-1] == 0.0:
                    vals[-1] = adj
                else:
                    vals.append(adj)
                if verbose and abs(adj) > 0.1:
                    print(f"  {repr(prev_char)}→{repr(next_char)}: {adj}", file=sys.stderr)
            # Don't update prev_char — adjustment sits between two chars
        i += 1


if __name__ == '__main__':
    main()
