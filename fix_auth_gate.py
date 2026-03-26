import codecs
import re

def main():
    with codecs.open('mcdm_app.py', 'r', 'utf-8') as f:
        lines = f.read().splitlines()

    start_idx = -1
    end_idx = -1
    for i, line in enumerate(lines):
        if line.strip() == 'st.markdown(' and '_vid = tt("Tanıtım Videosu", "Demo Video")' in lines[i-1]:
            start_idx = i
        if start_idx != -1 and line.strip() == ')':
            if 'unsafe_allow_html=True' in lines[i-1]:
                end_idx = i
                break

    if start_idx != -1 and end_idx != -1:
        new_block = [
            '    st.markdown(',
            '        f"""<div style="background-color: #020b18; background-image: radial-gradient(1px 1px at 5% 10%, rgba(255,255,255,0.95) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 12% 22%, rgba(255,255,255,0.8) 0%, transparent 100%), radial-gradient(1px 1px at 18% 5%, rgba(255,255,255,0.9) 0%, transparent 100%), radial-gradient(2px 2px at 25% 35%, rgba(255,255,255,0.6) 0%, transparent 100%), radial-gradient(1px 1px at 30% 15%, rgba(200,220,255,0.9) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 38% 48%, rgba(255,255,255,0.7) 0%, transparent 100%), radial-gradient(1px 1px at 43% 8%, rgba(255,255,255,0.85) 0%, transparent 100%), radial-gradient(2px 2px at 50% 28%, rgba(180,200,255,0.8) 0%, transparent 100%), radial-gradient(1px 1px at 55% 60%, rgba(255,255,255,0.7) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 62% 18%, rgba(255,255,255,0.9) 0%, transparent 100%), radial-gradient(1px 1px at 68% 42%, rgba(255,255,255,0.75) 0%, transparent 100%), radial-gradient(2px 2px at 74% 12%, rgba(200,215,255,0.85) 0%, transparent 100%), radial-gradient(1px 1px at 80% 55%, rgba(255,255,255,0.8) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 85% 30%, rgba(255,255,255,0.65) 0%, transparent 100%), radial-gradient(1px 1px at 90% 8%, rgba(255,255,255,0.9) 0%, transparent 100%), radial-gradient(2px 2px at 95% 45%, rgba(180,210,255,0.7) 0%, transparent 100%), radial-gradient(1px 1px at 8% 72%, rgba(255,255,255,0.6) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 15% 85%, rgba(255,255,255,0.75) 0%, transparent 100%), radial-gradient(1px 1px at 22% 65%, rgba(255,255,255,0.85) 0%, transparent 100%), radial-gradient(1px 1px at 35% 78%, rgba(200,220,255,0.7) 0%, transparent 100%), radial-gradient(2px 2px at 48% 88%, rgba(255,255,255,0.6) 0%, transparent 100%), radial-gradient(1px 1px at 58% 75%, rgba(255,255,255,0.8) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 70% 82%, rgba(255,255,255,0.7) 0%, transparent 100%), radial-gradient(1px 1px at 78% 68%, rgba(180,200,255,0.85) 0%, transparent 100%), radial-gradient(1px 1px at 88% 90%, rgba(255,255,255,0.65) 0%, transparent 100%), radial-gradient(1.5px 1.5px at 93% 72%, rgba(255,255,255,0.8) 0%, transparent 100%), radial-gradient(ellipse at 70% 20%, rgba(20,50,100,0.5) 0%, transparent 55%), radial-gradient(ellipse at 15% 60%, rgba(10,30,70,0.4) 0%, transparent 45%), linear-gradient(180deg, #020810 0%, #040e1f 40%, #061228 100%); border-radius: 16px; padding: 2.8rem 2.5rem 2.4rem 2.5rem; margin-bottom: 1.5rem; position: relative; overflow: hidden;">',
            '<div><div style="font-size:0.78rem;font-weight:500;letter-spacing:0.12em;color:#60A5FA;text-transform:uppercase;margin-bottom:0.25rem;">✦ &nbsp; {_subtitle}</div>',
            '<div style="font-size:0.92rem;color:#94A3B8;font-style:italic;margin-bottom:0.7rem;letter-spacing:0.02em;">Prof. Dr. Ömer Faruk Rençber</div>',
            '<h1 style="font-size:2.5rem;font-weight:800;color:#FFFFFF;margin:0 0 0.7rem 0;line-height:1.15;text-shadow:0 0 30px rgba(96,165,250,0.3);">MCDM Toolbox</h1>',
            '<p style="font-size:1.02rem;color:#CBD5E1;max-width:680px;line-height:1.7;margin:0 0 1.5rem 0;">{_desc}</p>',
            '<div style="display:flex;flex-wrap:wrap;gap:0.4rem;margin-bottom:1.8rem;">{_badge_html}</div>',
            '<div style="display:flex;flex-wrap:wrap;gap:0.75rem;align-items:center;">',
            '<div style="background:#F59E0B;color:#1B1B1B;border-radius:10px;padding:0.65rem 1.3rem;font-size:0.93rem;font-weight:700;">🔐 {_cta}</div>',
            '<a href="https://youtu.be/jp4oih6_Nec" target="_blank" style="display:inline-block;background:#DC2626;color:#FFFFFF;text-decoration:none;border-radius:10px;padding:0.65rem 1.1rem;font-size:0.9rem;font-weight:600;">🎥 {_vid}</a>',
            '<a href="https://www.instagram.com/mcdm_dss/" target="_blank" style="display:inline-block;background:#C13584;color:#FFFFFF;text-decoration:none;border-radius:10px;padding:0.65rem 1.1rem;font-size:0.9rem;font-weight:600;">📸 @mcdm_dss</a>',
            '</div></div></div>""",',
            '        unsafe_allow_html=True,',
            '    )'
        ]
        
        del lines[start_idx:end_idx+1]
        for n_idx, n_line in enumerate(new_block):
            lines.insert(start_idx + n_idx, n_line)

    content = '\\n'.join(lines)
    
    # Rerun logic to redirect to auth gate
    content = content.replace('access.logout_user()\\n        st.stop()', 'access.logout_user()\\n        st.rerun()')

    # Remove the privacy notice
    content = re.sub(r'^[ \\t]*st\.caption\(tt\(auth_settings\.privacy_notice_tr.*?$', '', content, flags=re.MULTILINE)

    with codecs.open('mcdm_app.py', 'w', 'utf-8') as f:
        f.write(content + '\\n')
    print("Corrections applied.")

if __name__ == '__main__':
    main()
