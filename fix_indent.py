import codecs

def main():
    with codecs.open("mcdm_app.py", "r", "utf-8") as f:
        lines = f.readlines()
        
    start_idx = -1
    end_idx = -1
    
    for i, line in enumerate(lines):
        if "st.markdown(f\"<div class='sb-section-label'>🧹 {tt('Veri Ön İşleme'" in line:
            start_idx = i
        elif "_step_data_done = is_data_loaded" in line and start_idx != -1:
            end_idx = i
            break
            
    if start_idx != -1 and end_idx != -1:
        # Replace the st.markdown line with st.expander
        # The original line has 12 spaces indentation
        lines[start_idx] = '            with st.expander(f"🧹 {tt(\'Veri Ön İşleme\', \'Data Preprocessing\')}", expanded=False):\n'
        
        # Indent the block between start_idx + 1 and end_idx
        for i in range(start_idx + 1, end_idx):
            # Only indent lines that are not completely empty
            if lines[i].strip():
                lines[i] = "    " + lines[i]
                
        with codecs.open("mcdm_app.py", "w", "utf-8") as f:
            f.writelines(lines)
        print(f"Successfully modified from line {start_idx} to {end_idx}")
    else:
        print("Could not find the target block.")

if __name__ == "__main__":
    main()
