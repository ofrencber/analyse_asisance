import subprocess
import codecs
import re

def fix_syntax():
    for i in range(1000):  # max 1000 fixes
        result = subprocess.run(['python', '-m', 'py_compile', 'mcdm_app.py'], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"Compilation successful after {i} fixes!")
            return True
            
        error = result.stderr
        # Parse line number
        match = re.search(r'line (\d+)', error)
        if not match:
            print("Cannot parse line from error:", error)
            return False
            
        line_num = int(match.group(1))
        
        with codecs.open('mcdm_app.py', 'r', 'utf-8') as f:
            lines = f.readlines()
            
        # Join line_num - 1 and line_num with a literal '\n'
        l1 = lines[line_num - 1].rstrip('\n')
        l2 = lines[line_num]
        
        lines[line_num - 1] = l1 + '\\n' + l2
        del lines[line_num]
        
        with codecs.open('mcdm_app.py', 'w', 'utf-8') as f:
            f.writelines(lines)
            
    print("Failed to converge after 1000 tries.")
    return False

if __name__ == '__main__':
    fix_syntax()
