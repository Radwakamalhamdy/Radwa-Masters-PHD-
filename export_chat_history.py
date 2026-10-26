# -*- coding: utf-8 -*-
"""
export_chat_history.py
Exports the entire conversation history from transcript_full.jsonl to a beautiful, clean Markdown (.md) file.
"""
import sys
import io
import json
import re
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

log_file = r'C:\Users\radwa kamal\.gemini\antigravity-ide\brain\bc34c39f-48ff-4509-8701-7daeb648b878\.system_generated\logs\transcript_full.jsonl'
output_md_1 = r'c:\Users\radwa kamal\Downloads\Masters & PHD Radwa\chat_history.md'
output_md_2 = r'c:\Users\radwa kamal\Downloads\Masters & PHD Radwa\Chat_History_Radwa_Scholarships.md'

turns = []
with open(log_file, 'r', encoding='utf-8') as f:
    for line in f:
        d = json.loads(line)
        t = d.get('type')
        c = d.get('content')
        if t == 'USER_INPUT' and c:
            # Clean up <USER_REQUEST> tags and any attached metadata
            user_text = c
            # Extract content inside <USER_REQUEST>...</USER_REQUEST>
            match = re.search(r'<USER_REQUEST>(.*?)</USER_REQUEST>', user_text, re.DOTALL)
            if match:
                clean_user_text = match.group(1).strip()
            else:
                clean_user_text = re.sub(r'<[^>]+>', '', user_text).strip()
            # Mask any GitHub Personal Access Tokens for security
            clean_user_text = re.sub(r'ghp_[A-Za-z0-9_]+', '[PROTECTED_GITHUB_ACCESS_TOKEN]', clean_user_text)
            turns.append(('User', clean_user_text))
        elif t == 'PLANNER_RESPONSE' and c:
            clean_assistant_text = re.sub(r'ghp_[A-Za-z0-9_]+', '[PROTECTED_GITHUB_ACCESS_TOKEN]', c.strip())
            turns.append(('Assistant', clean_assistant_text))

# Prepare Markdown document
md_content = []
md_content.append("# 📜 سجل المحادثة الكاملة والتطوير | Full Chat History & Development Log")
md_content.append("## 🎓 مشروع بوابة وقاعدة بيانات منح الماجستير والدكتوراه ومطابقة الـ CV")
md_content.append(f"**المستخدم / الباحثة:** رضوى كمال (Radwa Kamal, BSc Biotechnology - MSA & Greenwich)")
md_content.append(f"**المساعد الذكي:** Antigravity AI Assistant (Google DeepMind)")
md_content.append(f"**تاريخ التوثيق:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
md_content.append(f"**رابط المستودع على GitHub:** [https://github.com/Radwakamalhamdy/Radwa-Masters-PHD-](https://github.com/Radwakamalhamdy/Radwa-Masters-PHD-)")
md_content.append(f"**رابط الموقع الحي المباشر (GitHub Pages):** [https://radwakamalhamdy.github.io/Radwa-Masters-PHD-/](https://radwakamalhamdy.github.io/Radwa-Masters-PHD-/)")
md_content.append("\n---\n")

md_content.append("### 📑 فهرس محاور المحادثة (Table of Contents)")
md_content.append("1. **المحور الأول:** الفهم الشامل والتحليلي لمحتويات وهيكلية المجلد وخط أنابيب البيانات.")
md_content.append("2. **المحور الثاني:** إضافة المتطلبات التفصيلية لجميع المنح الـ 148 وحساب «طرح الـ CV» ونسب التوافق والجاهزية.")
md_content.append("3. **المحور الثالث:** التحقق من حالة النشر على GitHub والتأكد من عمل الرابط الحي (Live URL).")
md_content.append("4. **المحور الرابع:** إرشادات استخراج Personal Access Token من GitHub لإتمام الرفع.")
md_content.append("5. **المحور الخامس:** إتمام عملية الـ Git Push بنجاح ونشر التحديثات المكتملة على GitHub Pages.")
md_content.append("6. **المحور السادس:** توثيق وحفظ سجل المحادثة الكامل كملف Markdown.")
md_content.append("\n---\n")

turn_count = 1
for role, text in turns:
    if role == 'User':
        md_content.append(f"## 👤 الطلب #{turn_count} — رضوى كمال (User)")
        md_content.append("```text")
        md_content.append(text)
        md_content.append("```\n")
    else:
        md_content.append(f"## 🤖 الرد #{turn_count} — المساعد الذكي (Antigravity Assistant)")
        md_content.append(text)
        md_content.append("\n---\n")
        turn_count += 1

# Add final response section for Turn 6
if turns and turns[-1][0] == 'User':
    md_content.append(f"## 🤖 الرد #{turn_count} — المساعد الذكي (Antigravity Assistant)")
    md_content.append("تم حفظ وتوثيق سجل المحادثة الكاملة بنجاح في ملف Markdown منسق واحترافي داخل مجلد المشروع:")
    md_content.append("- `chat_history.md`")
    md_content.append("- `Chat_History_Radwa_Scholarships.md`")
    md_content.append("\nيحتوي الملف على جميع الحوارات، الأكواد، التحليلات، وروابط النشر الرسمية.")
    md_content.append("\n---\n")

full_md_text = "\n".join(md_content)

with open(output_md_1, 'w', encoding='utf-8') as f:
    f.write(full_md_text)
print(f"Saved: {output_md_1}")

with open(output_md_2, 'w', encoding='utf-8') as f:
    f.write(full_md_text)
print(f"Saved: {output_md_2}")
