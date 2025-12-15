# -*- coding: utf-8 -*-
"""
VBA Code Obfuscator - Làm rối code VBA để bảo vệ source code
Chạy trước khi build để obfuscate các module quan trọng
"""
from pathlib import Path
import re
import random
import string

BASE_DIR = Path(__file__).resolve().parent
SOURCE_DIR = BASE_DIR / "extracted_clean"
OBFUSCATED_DIR = BASE_DIR / "extracted_obfuscated"

# Modules cần obfuscate (không obfuscate UI forms để dễ maintain)
MODULES_TO_OBFUSCATE = [
    "modLicenseAudit.bas",
    "modAutoUpdate.bas",
    # Thêm các modules khác nếu cần
]

# VBA keywords không được đổi tên
VBA_KEYWORDS = {
    'Sub', 'End', 'Function', 'If', 'Then', 'Else', 'ElseIf', 'Select', 'Case',
    'For', 'Next', 'Do', 'Loop', 'While', 'Wend', 'With', 'Exit', 'GoTo',
    'On', 'Error', 'Resume', 'Call', 'Dim', 'As', 'Integer', 'String', 'Long',
    'Double', 'Boolean', 'Variant', 'Object', 'Date', 'Public', 'Private',
    'Const', 'Type', 'Enum', 'Option', 'Explicit', 'To', 'Step', 'Each',
    'In', 'Is', 'Not', 'And', 'Or', 'Xor', 'Like', 'Mod', 'New', 'Set',
    'Let', 'Get', 'ByVal', 'ByRef', 'Optional', 'ParamArray', 'Preserve',
    'ReDim', 'Erase', 'LBound', 'UBound', 'Array', 'True', 'False', 'Nothing',
    'Null', 'Empty', 'VbCrLf', 'vbCrLf', 'vbCr', 'vbLf', 'vbTab'
}

class VBAObfuscator:
    def __init__(self):
        self.var_map = {}  # Original name -> Obfuscated name
        self.counter = 0

    def generate_random_name(self, prefix=''):
        """Generate random variable name"""
        self.counter += 1
        # Tạo tên ngắn gọn: a1, a2, b1, b2, ...
        if prefix:
            return f"{prefix}_{self.counter}"
        else:
            letter = chr(97 + (self.counter // 100) % 26)  # a-z
            return f"{letter}{self.counter}"

    def extract_variables(self, code):
        """Extract all variable names from code"""
        variables = set()

        # Tìm Dim declarations
        dim_pattern = r'Dim\s+(\w+)\s+As'
        for match in re.finditer(dim_pattern, code, re.IGNORECASE):
            var_name = match.group(1)
            if var_name not in VBA_KEYWORDS:
                variables.add(var_name)

        # Tìm Function/Sub names
        func_pattern = r'(?:Public|Private|Friend)?\s*(?:Sub|Function)\s+(\w+)'
        for match in re.finditer(func_pattern, code, re.IGNORECASE):
            func_name = match.group(1)
            if func_name not in VBA_KEYWORDS and not func_name.startswith('Workbook_'):
                variables.add(func_name)

        # Tìm Const declarations
        const_pattern = r'Const\s+(\w+)\s*='
        for match in re.finditer(const_pattern, code, re.IGNORECASE):
            const_name = match.group(1)
            if const_name not in VBA_KEYWORDS:
                variables.add(const_name)

        return variables

    def obfuscate_code(self, code, keep_public=True):
        """Obfuscate VBA code"""

        # 1. Extract variables
        variables = self.extract_variables(code)

        # 2. Create mapping for each variable
        for var in variables:
            # Giữ nguyên Public functions để có thể gọi từ bên ngoài
            if keep_public and re.search(rf'Public\s+(?:Sub|Function)\s+{var}\b', code, re.IGNORECASE):
                continue

            if var not in self.var_map:
                self.var_map[var] = self.generate_random_name()

        # 3. Replace variables in code
        obfuscated = code
        for original, obfuscated_name in self.var_map.items():
            # Replace with word boundaries to avoid partial replacements
            obfuscated = re.sub(
                rf'\b{original}\b',
                obfuscated_name,
                obfuscated
            )

        # 4. Remove comments (optional - giữ lại để dễ debug)
        # obfuscated = re.sub(r"'[^\r\n]*", '', obfuscated)

        # 5. Remove blank lines
        lines = obfuscated.split('\n')
        lines = [line for line in lines if line.strip()]
        obfuscated = '\n'.join(lines)

        return obfuscated

    def obfuscate_string_literals(self, code):
        """Obfuscate string literals to make them harder to read"""
        def encode_string(match):
            text = match.group(1)
            # Convert to Chr() calls
            chars = [f'Chr({ord(c)})' for c in text]
            return ' & '.join(chars)

        # Chỉ obfuscate strings dài hơn 10 ký tự (tránh ảnh hưởng UI)
        pattern = r'"([^"]{10,})"'
        return re.sub(pattern, lambda m: encode_string(m), code)


def obfuscate_modules():
    """Obfuscate specified VBA modules"""
    if not SOURCE_DIR.exists():
        print(f"ERROR: Source directory not found: {SOURCE_DIR}")
        return

    # Create obfuscated directory
    OBFUSCATED_DIR.mkdir(exist_ok=True)

    obfuscator = VBAObfuscator()

    print("VBA Code Obfuscator")
    print("=" * 50)

    for module_name in MODULES_TO_OBFUSCATE:
        source_file = SOURCE_DIR / module_name
        if not source_file.exists():
            print(f"⚠️  Module not found: {module_name}")
            continue

        print(f"\n📝 Obfuscating: {module_name}")

        # Read original code
        try:
            code = source_file.read_text(encoding='utf-8')
        except UnicodeDecodeError:
            code = source_file.read_text(encoding='cp1252')

        # Separate header (Attribute lines) from code
        lines = code.split('\n')
        header_lines = []
        code_lines = []
        in_header = True

        for line in lines:
            stripped = line.strip()
            if in_header and (stripped.startswith('VERSION') or
                             stripped.startswith('BEGIN') or
                             stripped.startswith('END') or
                             stripped.startswith('Attribute') or
                             stripped == ''):
                header_lines.append(line)
            else:
                in_header = False
                code_lines.append(line)

        header = '\n'.join(header_lines)
        code_only = '\n'.join(code_lines)

        # Obfuscate code
        obfuscated_code = obfuscator.obfuscate_code(code_only, keep_public=True)

        # Optional: obfuscate strings (có thể bỏ comment nếu muốn)
        # obfuscated_code = obfuscator.obfuscate_string_literals(obfuscated_code)

        # Combine header + obfuscated code
        final_code = header + '\n' + obfuscated_code

        # Save to obfuscated directory
        output_file = OBFUSCATED_DIR / module_name
        output_file.write_text(final_code, encoding='utf-8')

        print(f"✅ Saved to: {output_file}")
        print(f"   Variables obfuscated: {len(obfuscator.var_map)}")

    # Copy non-obfuscated modules
    print(f"\n📋 Copying non-obfuscated modules...")
    for source_file in SOURCE_DIR.glob("*.bas"):
        if source_file.name not in MODULES_TO_OBFUSCATE:
            output_file = OBFUSCATED_DIR / source_file.name
            output_file.write_bytes(source_file.read_bytes())

    for source_file in SOURCE_DIR.glob("*.cls"):
        output_file = OBFUSCATED_DIR / source_file.name
        output_file.write_bytes(source_file.read_bytes())

    for source_file in SOURCE_DIR.glob("*.frm"):
        output_file = OBFUSCATED_DIR / source_file.name
        output_file.write_bytes(source_file.read_bytes())
        # Copy frx files too
        frx_file = source_file.with_suffix('.frx')
        if frx_file.exists():
            output_frx = OBFUSCATED_DIR / frx_file.name
            output_frx.write_bytes(frx_file.read_bytes())

    print(f"\n✅ Done! Obfuscated modules saved to: {OBFUSCATED_DIR}")
    print(f"\n💡 Next step: Update rebuild_xlam.py to use 'extracted_obfuscated' folder")
    print(f"   Or manually copy obfuscated files to 'extracted_clean'")


if __name__ == "__main__":
    obfuscate_modules()
