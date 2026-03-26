import codecs

def main():
    with codecs.open("mcdm_app.py", "r", "utf-8") as f:
        content = f.read()
        
    auth_block_start = content.find("# ---------------------------------------------------------\n# AUTH GATE — Giriş zorunluysa login ekranı")
    if auth_block_start == -1:
        print("Auth block start not found.")
        return
        
    auth_block_end_marker = "        st.stop()"
    # Find the SECOND st.stop() after the start, which corresponds to the end of the name collection screen block.
    first_stop = content.find(auth_block_end_marker, auth_block_start)
    second_stop = content.find(auth_block_end_marker, first_stop + len(auth_block_end_marker))
    
    if second_stop == -1:
        print("Auth block end not found.")
        return
        
    # include the st.stop() line
    end_pos = content.find('\n', second_stop) + 1
    
    auth_block_str = content[auth_block_start:end_pos]
    
    # Remove the auth block from its original location
    content = content[:auth_block_start] + content[end_pos:]
    
    # Find where to insert it: right after _current_user = access.get_current_user()
    insert_marker = "_current_user = access.get_current_user()\n"
    insert_pos = content.find(insert_marker)
    if insert_pos == -1:
        print("Insert position not found.")
        return
        
    insert_pos += len(insert_marker)
    
    # Insert the auth block
    new_content = content[:insert_pos] + "\n" + auth_block_str + "\n" + content[insert_pos:]
    
    with codecs.open("mcdm_app.py", "w", "utf-8") as f:
        f.write(new_content)
        
    print("Successfully moved Auth Gate block to the top of the execution flow!")

if __name__ == "__main__":
    main()
