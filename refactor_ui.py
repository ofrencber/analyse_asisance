import re

def main():
    with open("mcdm_app.py", "r", encoding="utf-8") as f:
        lines = f.readlines()

    lines = [line.rstrip('\n') for line in lines]

    start_idx = -1
    end_idx = -1
    
    for i, line in enumerate(lines):
        if "with st.expander(_step1_label" in line and start_idx == -1:
            start_idx = i
        if 'st.session_state["sensitivity_sigma"] = float(sensitivity_sigma)' in line:
            end_idx = i

    if start_idx == -1 or end_idx == -1:
        print("Indices not found!")
        return

    new_lines = []
    new_lines.extend(lines[:start_idx])
    
    new_lines.append('_has_results = bool(st.session_state.get("last_results"))')
    new_lines.append('with st.expander(tt("⚙️ Analiz Kurulumu (1-2-3. Adımlar)", "⚙️ Analysis Setup (Steps 1-3)"), expanded=not _has_results):')

    for i in range(start_idx, end_idx + 1):
        line = lines[i]
        
        if line.startswith('with st.expander(_step1_label'):
            new_lines.append('    st.markdown(f"#### {_step1_label}")')
            new_lines.append('    with st.container():')
        elif line.startswith('with st.expander(_prep_label'):
            new_lines.append('    st.markdown(f"#### {_prep_label}")')
            new_lines.append('    with st.container():')
        elif line.startswith('with st.expander(_step3_title'):
            new_lines.append('    st.markdown(f"#### {_step3_title}")')
            new_lines.append('    with st.container():')
        elif line.startswith('st.divider()') and i > start_idx:
            # We also indent dividers
            new_lines.append('    ' + line)
        else:
            # Indent the existing line by 4 spaces (if it's not totally empty)
            if line.strip():
                new_lines.append('    ' + line)
            else:
                new_lines.append('')

    new_lines.extend(lines[end_idx+1:])

    with open("mcdm_app.py", "w", encoding="utf-8") as f:
        f.write("\n".join(new_lines) + "\n")
        
    print("Refactoring complete.")

if __name__ == "__main__":
    main()
