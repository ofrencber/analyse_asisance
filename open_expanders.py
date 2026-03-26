import sys

def process_file(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Replace exactly Tablolar and Şekiller
    content = content.replace(
        'with st.expander(f"📋 {tt(\'Tablolar\', \'Tables\')}", expanded=False):',
        'with st.expander(f"📋 {tt(\'Tablolar\', \'Tables\')}", expanded=True):'
    )
    content = content.replace(
        'with st.expander(f"📊 {tt(\'Şekiller\', \'Figures\')}", expanded=False):',
        'with st.expander(f"📊 {tt(\'Şekiller\', \'Figures\')}", expanded=True):'
    )
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
        
    print("Expanders updated.")

if __name__ == '__main__':
    process_file('mcdm_app.py')
