# -*- coding: utf-8 -*-
"""
generate_html_app.py
Generates the interactive single-page web portal for Radwa Kamal's Master's & PhD Scholarships Hub.
Includes:
- Detailed requirements for each scholarship
- CV Subtraction / Gap Analysis (ما يستوفيه ملف رضوى vs المتبقي لاستيفاء الشروط 100%)
- Match Rate & Readiness badges
- Instant live search & multi-filtering
- In-browser LocalStorage editing and persistence
- SheetJS instant Excel export with all new columns
"""
import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('curated_scholarships.json', 'r', encoding='utf-8') as f:
    scholarships_data = json.load(f)

print(f"Embedding {len(scholarships_data)} items into HTML...")

data_json_str = json.dumps(scholarships_data, ensure_ascii=False)

html_template = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>بوابة وقاعدة بيانات منح الماجستير والدكتوراه ومطابقة الـ CV | رضوى كمال</title>
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
    
    <!-- SheetJS for Instant Excel Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <style>
        :root {
            --bg-primary: #f8fafc;
            --bg-secondary: #ffffff;
            --bg-card: rgba(255, 255, 255, 0.96);
            --bg-card-hover: #f1f5f9;
            --border-color: #e2e8f0;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --primary: #059669;
            --primary-hover: #047857;
            --primary-light: #ecfdf5;
            --accent: #2563eb;
            --accent-light: #eff6ff;
            --warning: #d97706;
            --warning-light: #fffbeb;
            --purple: #7c3aed;
            --purple-light: #f5f3ff;
            --danger: #dc2626;
            --danger-light: #fef2f2;
            --shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.06), 0 2px 6px -1px rgba(0, 0, 0, 0.04);
            --radius-sm: 8px;
            --radius-md: 14px;
            --radius-lg: 20px;
            --transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }

        [data-theme="dark"] {
            --bg-primary: #0b0f19;
            --bg-secondary: #111827;
            --bg-card: rgba(17, 24, 39, 0.85);
            --bg-card-hover: rgba(31, 41, 55, 0.95);
            --border-color: rgba(255, 255, 255, 0.08);
            --text-main: #f9fafb;
            --text-muted: #9ca3af;
            --primary: #10b981;
            --primary-hover: #059669;
            --primary-light: rgba(16, 185, 129, 0.15);
            --accent: #3b82f6;
            --accent-light: rgba(59, 130, 246, 0.15);
            --warning: #f59e0b;
            --warning-light: rgba(245, 158, 11, 0.15);
            --purple: #8b5cf6;
            --purple-light: rgba(139, 92, 246, 0.15);
            --danger: #ef4444;
            --danger-light: rgba(239, 68, 68, 0.15);
            --shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Cairo', 'Plus Jakarta Sans', system-ui, sans-serif;
        }

        body {
            background-color: var(--bg-primary);
            color: var(--text-main);
            min-height: 100vh;
            line-height: 1.6;
            transition: background-color 0.3s ease, color 0.3s ease;
            padding-bottom: 80px;
        }

        /* Ambient Glow Accent */
        .ambient-glow {
            position: fixed;
            top: -200px;
            right: 15%;
            width: 600px;
            height: 600px;
            background: radial-gradient(circle, rgba(16, 185, 129, 0.12) 0%, rgba(37, 99, 235, 0.05) 50%, transparent 70%);
            filter: blur(80px);
            z-index: -1;
            pointer-events: none;
        }

        .container {
            max-width: 1750px;
            margin: 0 auto;
            padding: 24px 20px;
        }

        /* Top Header */
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
            padding: 24px 30px;
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow);
            margin-bottom: 24px;
        }

        .header-title-box {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .header-badge-icon {
            width: 58px;
            height: 58px;
            border-radius: 16px;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            box-shadow: 0 8px 16px rgba(5, 150, 105, 0.25);
            color: #fff;
        }

        .header-title h1 {
            font-size: 24px;
            font-weight: 800;
            color: var(--text-main);
            margin-bottom: 4px;
        }

        .header-title p {
            font-size: 13.5px;
            color: var(--text-muted);
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 12px;
            flex-wrap: wrap;
        }

        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 10px 18px;
            border-radius: var(--radius-md);
            font-size: 13.5px;
            font-weight: 700;
            cursor: pointer;
            border: none;
            transition: var(--transition);
            text-decoration: none;
        }

        .btn-primary {
            background: linear-gradient(135deg, var(--primary), var(--primary-hover));
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(5, 150, 105, 0.25);
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(5, 150, 105, 0.35);
        }

        .btn-accent {
            background: linear-gradient(135deg, var(--accent), #1d4ed8);
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
        }

        .btn-accent:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35);
        }

        .btn-secondary {
            background: var(--bg-secondary);
            color: var(--text-main);
            border: 1px solid var(--border-color);
        }

        .btn-secondary:hover {
            background: var(--bg-card-hover);
            transform: translateY(-2px);
        }

        .theme-toggle-btn {
            width: 44px;
            height: 44px;
            border-radius: var(--radius-md);
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            font-size: 18px;
            transition: var(--transition);
        }

        .theme-toggle-btn:hover {
            background: var(--bg-card-hover);
            transform: rotate(15deg);
        }

        /* Radwa Profile Banner */
        .profile-banner {
            background: linear-gradient(135deg, rgba(37, 99, 235, 0.08), rgba(16, 185, 129, 0.08));
            border: 1px solid rgba(37, 99, 235, 0.2);
            border-radius: var(--radius-lg);
            padding: 18px 24px;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 16px;
        }

        .profile-info {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .profile-avatar {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            background: linear-gradient(135deg, var(--accent), var(--purple));
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            font-weight: 800;
        }

        .profile-text h3 {
            font-size: 16px;
            font-weight: 800;
            color: var(--text-main);
            margin-bottom: 2px;
        }

        .profile-text p {
            font-size: 12.5px;
            color: var(--text-muted);
        }

        .profile-tags {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .ptag {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 11.5px;
            font-weight: 700;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-main);
        }

        /* Stats Grid */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .stat-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 18px 20px;
            box-shadow: var(--shadow);
            display: flex;
            align-items: center;
            gap: 16px;
            transition: var(--transition);
        }

        .stat-card:hover {
            transform: translateY(-3px);
            border-color: var(--primary);
        }

        .stat-icon {
            width: 48px;
            height: 48px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
        }

        .stat-icon.green { background: var(--primary-light); color: var(--primary); }
        .stat-icon.blue { background: var(--accent-light); color: var(--accent); }
        .stat-icon.amber { background: var(--warning-light); color: var(--warning); }
        .stat-icon.purple { background: var(--purple-light); color: var(--purple); }

        .stat-data h4 {
            font-size: 24px;
            font-weight: 800;
            line-height: 1.2;
            color: var(--text-main);
        }

        .stat-data p {
            font-size: 12.5px;
            color: var(--text-muted);
            font-weight: 600;
        }

        /* Controls Section */
        .controls-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            padding: 20px 24px;
            box-shadow: var(--shadow);
            margin-bottom: 24px;
        }

        .search-box {
            position: relative;
            margin-bottom: 16px;
        }

        .search-input {
            width: 100%;
            padding: 14px 48px 14px 20px;
            border-radius: var(--radius-md);
            border: 1.5px solid var(--border-color);
            background: var(--bg-secondary);
            color: var(--text-main);
            font-size: 14.5px;
            transition: var(--transition);
        }

        .search-input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 4px rgba(16, 185, 129, 0.15);
        }

        .search-icon {
            position: absolute;
            right: 18px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            font-size: 18px;
        }

        /* Filter Tabs */
        .filter-tabs {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 16px;
        }

        .tab-btn {
            padding: 8px 16px;
            border-radius: 30px;
            font-size: 12.5px;
            font-weight: 700;
            border: 1px solid var(--border-color);
            background: var(--bg-secondary);
            color: var(--text-muted);
            cursor: pointer;
            transition: var(--transition);
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .tab-btn:hover {
            color: var(--text-main);
            border-color: var(--text-muted);
        }

        .tab-btn.active {
            background: var(--primary);
            color: #ffffff;
            border-color: var(--primary);
            box-shadow: 0 4px 10px rgba(5, 150, 105, 0.25);
        }

        .tab-count {
            padding: 1px 7px;
            border-radius: 10px;
            font-size: 11px;
            background: rgba(0, 0, 0, 0.08);
        }

        .tab-btn.active .tab-count {
            background: rgba(255, 255, 255, 0.25);
            color: #ffffff;
        }

        /* Secondary Filters */
        .secondary-filters {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
            align-items: center;
        }

        .filter-select {
            padding: 8px 14px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            background: var(--bg-secondary);
            color: var(--text-main);
            font-size: 12.5px;
            font-weight: 600;
            cursor: pointer;
        }

        .filter-select:focus {
            outline: none;
            border-color: var(--accent);
        }

        /* Edit Notice Bar */
        .edit-notice-bar {
            background: var(--warning-light);
            border: 1px solid var(--warning);
            border-radius: var(--radius-md);
            padding: 12px 20px;
            margin-bottom: 20px;
            display: none;
            align-items: center;
            justify-content: space-between;
            animation: fadeIn 0.3s ease;
        }

        .edit-notice-bar.active {
            display: flex;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        /* Table Card & Responsive Table */
        .table-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow);
            overflow: hidden;
        }

        .table-responsive {
            width: 100%;
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: right;
            font-size: 13px;
        }

        th {
            background: var(--bg-card-hover);
            color: var(--text-muted);
            padding: 14px 16px;
            font-weight: 700;
            font-size: 12.5px;
            border-bottom: 1px solid var(--border-color);
            white-space: nowrap;
        }

        td {
            padding: 16px;
            border-bottom: 1px solid var(--border-color);
            vertical-align: top;
            transition: background 0.2s ease;
        }

        tr:hover td {
            background: var(--bg-card-hover);
        }

        /* Scholarship Cell */
        .scholarship-title {
            font-weight: 800;
            font-size: 14.5px;
            color: var(--text-main);
            margin-bottom: 4px;
        }

        .scholarship-title-en {
            font-size: 12px;
            color: var(--text-muted);
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-weight: 600;
            margin-bottom: 8px;
        }

        .tags-row {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            margin-top: 6px;
        }

        .badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
        }

        .badge-africa { background: #fef3c7; color: #92400e; }
        .badge-gov { background: #dbeafe; color: #1e40af; }
        .badge-biotech { background: #dcfce7; color: #166534; }
        .badge-uni { background: #f3e8ff; color: #6b21a8; }
        .badge-dir { background: #fce7f3; color: #9d174d; }

        .badge-degree {
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            color: var(--text-main);
        }

        /* Match Rate Pill */
        .match-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 800;
            color: #ffffff;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 6px;
        }

        .match-95 { background: linear-gradient(135deg, #059669, #10b981); }
        .match-90 { background: linear-gradient(135deg, #2563eb, #3b82f6); }
        .match-85 { background: linear-gradient(135deg, #d97706, #f59e0b); }
        .match-80 { background: linear-gradient(135deg, #ea580c, #f97316); }

        .readiness-text {
            font-size: 11.5px;
            color: var(--text-muted);
            font-weight: 600;
            display: block;
            line-height: 1.4;
        }

        /* Requirements Box */
        .req-box {
            font-size: 12px;
            line-height: 1.6;
            color: var(--text-main);
            background: var(--bg-primary);
            padding: 10px 12px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            white-space: pre-line;
            max-width: 380px;
        }

        /* CV Gap Analysis Box */
        .gap-box {
            font-size: 12px;
            line-height: 1.6;
            background: linear-gradient(180deg, rgba(236, 253, 245, 0.4), rgba(239, 246, 255, 0.4));
            border: 1px solid rgba(16, 185, 129, 0.25);
            border-radius: var(--radius-sm);
            padding: 10px 12px;
            white-space: pre-line;
            color: var(--text-main);
            max-width: 440px;
        }

        [data-theme="dark"] .gap-box {
            background: rgba(16, 185, 129, 0.05);
            border-color: rgba(16, 185, 129, 0.2);
        }

        .link-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: var(--accent);
            text-decoration: none;
            font-weight: 700;
            font-size: 11.5px;
            padding: 4px 10px;
            border-radius: 6px;
            background: var(--accent-light);
            border: 1px solid rgba(37, 99, 235, 0.2);
            transition: var(--transition);
        }

        .link-pill:hover {
            background: var(--accent);
            color: white;
            transform: translateY(-1px);
        }

        .date-box {
            font-size: 11.5px;
            font-weight: 600;
            color: var(--text-main);
            display: block;
            margin-bottom: 4px;
        }

        .date-deadline {
            color: #dc2626;
            font-weight: 700;
        }

        [data-theme="dark"] .date-deadline {
            color: #f87171;
        }

        .action-btns-cell {
            display: flex;
            align-items: center;
            gap: 6px;
            justify-content: center;
        }

        .icon-btn {
            width: 32px;
            height: 32px;
            border-radius: 8px;
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 13px;
            transition: var(--transition);
        }

        .icon-btn:hover {
            color: var(--primary);
            border-color: var(--primary);
            transform: scale(1.08);
        }

        .icon-btn.delete-btn:hover {
            color: var(--danger);
            border-color: var(--danger);
        }

        /* Detail Modal */
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.65);
            backdrop-filter: blur(8px);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 999;
            opacity: 0;
            visibility: hidden;
            transition: all 0.3s ease;
            padding: 16px;
        }

        .modal-overlay.active {
            opacity: 1;
            visibility: visible;
        }

        .modal-box {
            background: var(--bg-secondary);
            border-radius: var(--radius-lg);
            width: 100%;
            max-width: 820px;
            max-height: 90vh;
            overflow-y: auto;
            border: 1px solid var(--border-color);
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
            padding: 28px;
        }

        .modal-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 14px;
        }

        .modal-header h3 {
            font-size: 19px;
            font-weight: 800;
        }

        .close-modal-btn {
            background: none;
            border: none;
            font-size: 24px;
            color: var(--text-muted);
            cursor: pointer;
        }

        .form-group {
            margin-bottom: 16px;
        }

        .form-group label {
            display: block;
            font-size: 12.5px;
            font-weight: 700;
            margin-bottom: 6px;
            color: var(--text-main);
        }

        .form-control {
            width: 100%;
            padding: 10px 14px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            background: var(--bg-primary);
            color: var(--text-main);
            font-size: 13px;
        }

        .form-control:focus {
            outline: none;
            border-color: var(--primary);
        }

        /* Toast */
        .toast {
            position: fixed;
            bottom: 30px;
            left: 30px;
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            padding: 14px 22px;
            border-radius: var(--radius-md);
            box-shadow: var(--shadow);
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 13.5px;
            font-weight: 700;
            z-index: 1000;
            transform: translateY(100px);
            opacity: 0;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .toast.show {
            transform: translateY(0);
            opacity: 1;
        }
    </style>
</head>
<body>
    <div class="ambient-glow"></div>

    <div class="container">
        <!-- Header -->
        <header>
            <div class="header-title-box">
                <div class="header-badge-icon">🎓</div>
                <div class="header-title">
                    <h1>بوابة وقاعدة بيانات منح الماجستير والدكتوراه ومطابقة الـ CV</h1>
                    <p>الموسوعة الشاملة لفرص الدراسات العليا لطلاب مصر وأفريقيا وتخصصات التكنولوجيا الحيوية والعلوم الطبية</p>
                </div>
            </div>
            <div class="header-actions">
                <button class="btn btn-primary" onclick="exportToExcel()">
                    📊 تصدير Excel بالمطابقة
                </button>
                <button class="btn btn-accent" onclick="openAddModal()">
                    ➕ إضافة منحة جديدة
                </button>
                <button class="btn btn-secondary" onclick="resetToOriginal()" title="استعادة النسخة الأصلية">
                    🔄 استعادة الأصل
                </button>
                <button class="theme-toggle-btn" onclick="toggleTheme()" title="تبديل الوضع الليلي / الفاتح">
                    <span id="themeIcon">🌙</span>
                </button>
            </div>
        </header>

        <!-- Profile Matching Banner -->
        <div class="profile-banner">
            <div class="profile-info">
                <div class="profile-avatar">RK</div>
                <div class="profile-text">
                    <h3>الملف الأكاديمي المرجعي للمطابقة: رضوى كمال (Radwa Kamal, BSc)</h3>
                    <p>بكالوريوس تكنولوجيا حيوية (MSA: 3.53 / جرينتش بريطانيا: 3.6 First Class) • ورقة بحثية محكمة + 10 بوسترات دولية • مؤسس شركتين ناشئتين • باحثة بمعمل النماذج الأولية</p>
                </div>
            </div>
            <div class="profile-tags">
                <span class="ptag">🧬 Biotechnology & Genetics</span>
                <span class="ptag">🤖 AI & Medical Robotics</span>
                <span class="ptag">🏆 First Class Honours</span>
                <span class="ptag">🗣️ English C1 / Dual Degree</span>
            </div>
        </div>

        <!-- Stats Grid -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon green">📊</div>
                <div class="stat-data">
                    <h4 id="statTotal">148</h4>
                    <p>إجمالي الفرص والمنظمات</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon green">⚡</div>
                <div class="stat-data">
                    <h4 id="statMatch95">42</h4>
                    <p>جاهزية فائقة (95%+ تنافسية فورية)</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon blue">🌍</div>
                <div class="stat-data">
                    <h4 id="statAfrica">26</h4>
                    <p>أفريقيا والدول النامية</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon purple">🧬</div>
                <div class="stat-data">
                    <h4 id="statBiotech">8</h4>
                    <p>معاهد البيوتكنولوجي المتخصصة</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon amber">💼</div>
                <div class="stat-data">
                    <h4 id="statPhd">31</h4>
                    <p>وظائف دكتوراه ممولة وأدلة بحث</p>
                </div>
            </div>
        </div>

        <!-- Controls Card -->
        <div class="controls-card">
            <!-- Live Search -->
            <div class="search-box">
                <span class="search-icon">🔍</span>
                <input type="text" id="searchInput" class="search-input" placeholder="ابحثي باسم المنحة، الدولة، التخصص، شروط الآيلتس، أو ما ينقصك في الـ CV..." oninput="handleSearch()">
            </div>

            <!-- Filter Tabs -->
            <div class="filter-tabs">
                <button class="tab-btn active" onclick="setCategory('all', this)">
                    كل المنح <span class="tab-count" id="countAll">148</span>
                </button>
                <button class="tab-btn" onclick="setCategory('أفريقيا والدول النامية', this)">
                    🌍 أفريقيا والدول النامية <span class="tab-count" id="countAfrica">26</span>
                </button>
                <button class="tab-btn" onclick="setCategory('المنح الحكومية والدولية الكبرى', this)">
                    🏛️ المنح الحكومية العالمية <span class="tab-count" id="countGov">52</span>
                </button>
                <button class="tab-btn" onclick="setCategory('مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية', this)">
                    🧬 معاهد البيوتكنولوجي والطب <span class="tab-count" id="countBiotechTab">8</span>
                </button>
                <button class="tab-btn" onclick="setCategory('منح الجامعات العالمية والمؤسسات المرموقة', this)">
                    🎓 الجامعات العالمية الكبرى <span class="tab-count" id="countUni">31</span>
                </button>
                <button class="tab-btn" onclick="setCategory('محركات البحث وأدلة وظائف الدكتوراه', this)">
                    💼 وظائف وأدلة الدكتوراه <span class="tab-count" id="countDir">31</span>
                </button>
            </div>

            <!-- Secondary Filters -->
            <div class="secondary-filters">
                <div>
                    <label style="font-size: 12px; font-weight: 700; color: var(--text-muted); margin-left: 6px;">المرحلة الدراسية:</label>
                    <select id="degreeFilter" class="filter-select" onchange="applyFilters()">
                        <option value="all">جميع المراحل</option>
                        <option value="ماجستير">ماجستير (Master's)</option>
                        <option value="دكتوراه">دكتوراه (PhD)</option>
                    </select>
                </div>

                <div>
                    <label style="font-size: 12px; font-weight: 700; color: var(--text-muted); margin-left: 6px;">نسبة التوافق مع ملفك:</label>
                    <select id="matchFilter" class="filter-select" onchange="applyFilters()">
                        <option value="all">جميع نسب الجاهزية</option>
                        <option value="95">جاهزية فائقة (95%+ تقديم فوري)</option>
                        <option value="90">مؤهلة بقوة (90%+ ينقص آيلتس فقط)</option>
                        <option value="85">مؤهلة (85%+ تتطلب مقترح/مشرف)</option>
                    </select>
                </div>

                <div>
                    <label style="font-size: 12px; font-weight: 700; color: var(--text-muted); margin-left: 6px;">نوع التمويل:</label>
                    <select id="fundingFilter" class="filter-select" onchange="applyFilters()">
                        <option value="all">كل أنواع التمويل</option>
                        <option value="تمويل كامل">تمويل كامل 100%</option>
                        <option value="وظيفة">وظيفة بحثية براتب</option>
                    </select>
                </div>
            </div>
        </div>

        <!-- Edit Notice Bar -->
        <div class="edit-notice-bar" id="editNoticeBar">
            <div style="display: flex; align-items: center; gap: 8px;">
                <span>✏️</span>
                <span style="font-size: 13.5px; font-weight: 700; color: #92400e;">لقد قمتِ بإجراء تعديلات في الجدول. لا تنسي حفظ التغييرات محلياً!</span>
            </div>
            <div style="display: flex; gap: 8px;">
                <button class="btn btn-primary" style="padding: 6px 14px; font-size: 12px;" onclick="saveToStorage(true)">💾 حفظ التعديلات الآن</button>
            </div>
        </div>

        <!-- Table Card -->
        <div class="table-card">
            <div class="table-responsive">
                <table id="scholarshipsTable">
                    <thead>
                        <tr>
                            <th style="width: 40px; text-align: center;">#</th>
                            <th style="width: 260px;">المنحة والتصنيف</th>
                            <th style="width: 170px; text-align: center;">نسبة التوافق مع الـ CV</th>
                            <th style="width: 360px;">المتطلبات التفصيلية والشروط</th>
                            <th style="min-width: 380px;">طرح الـ CV: المتبقي لاستيفاء الشروط 100%</th>
                            <th style="width: 170px;">مواعيد التقديم والتمويل</th>
                            <th style="width: 80px; text-align: center;">إجراءات</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Populated via Javascript -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Add Scholarship Modal -->
    <div class="modal-overlay" id="addModal">
        <div class="modal-box">
            <div class="modal-header">
                <h3>إضافة فرصة أو منظمة جديدة</h3>
                <button class="close-modal-btn" onclick="closeAddModal()">&times;</button>
            </div>
            <form id="addForm" onsubmit="handleAddNew(event)">
                <div class="form-group">
                    <label>اسم المنحة / المنظمة *</label>
                    <input type="text" id="newName" class="form-control" placeholder="مثال: منحة الحكومة الكندية فانير (Vanier CGS)" required>
                </div>
                <div class="form-group">
                    <label>التصنيف والنطاق *</label>
                    <select id="newCategory" class="form-control" required>
                        <option value="أفريقيا والدول النامية">أفريقيا والدول النامية</option>
                        <option value="المنح الحكومية والدولية الكبرى">المنح الحكومية والدولية الكبرى</option>
                        <option value="مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية">مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية</option>
                        <option value="منح الجامعات العالمية والمؤسسات المرموقة">منح الجامعات العالمية والمؤسسات المرموقة</option>
                        <option value="محركات البحث وأدلة وظائف الدكتوراه">محركات البحث وأدلة وظائف الدكتوراه</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>المرحلة الدراسية *</label>
                    <select id="newDegree" class="form-control" required>
                        <option value="ماجستير">ماجستير</option>
                        <option value="دكتوراه">دكتوراه</option>
                        <option value="ماجستير ودكتوراه">ماجستير ودكتوراه</option>
                        <option value="وظيفة بحثية ممولة">وظيفة بحثية ممولة</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>الرابط الرسمي المباشر *</label>
                    <input type="url" id="newLink" class="form-control" placeholder="https://..." required>
                </div>
                <div class="form-group">
                    <label>المتطلبات التفصيلية والشروط *</label>
                    <textarea id="newReqs" class="form-control" rows="3" placeholder="المؤهل المطلوب، شرط السن، اللغة، التخصصات..." required></textarea>
                </div>
                <div class="form-group">
                    <label>طرح الـ CV: المتبقي لاستيفاء الشروط 100%</label>
                    <textarea id="newGap" class="form-control" rows="3" placeholder="ما يستوفيه ملفك وما ينقصك للتقديم..."></textarea>
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
                    <div class="form-group">
                        <label>ميعاد بداية التقديم</label>
                        <input type="text" id="newStart" class="form-control" placeholder="مثال: أكتوبر سنوياً">
                    </div>
                    <div class="form-group">
                        <label>آخر ميعاد للتقديم (Deadline)</label>
                        <input type="text" id="newDeadline" class="form-control" placeholder="مثال: 15 يناير">
                    </div>
                </div>
                <div class="form-group">
                    <label>نوع التمويل</label>
                    <input type="text" id="newNotes" class="form-control" placeholder="مثال: تمويل كامل 100% يشمل السكن والراتب">
                </div>
                <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px;">
                    <button type="button" class="btn btn-secondary" onclick="closeAddModal()">إلغاء</button>
                    <button type="submit" class="btn btn-primary">حفظ الإضافة</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Quick View / Full Modal -->
    <div class="modal-overlay" id="quickViewModal">
        <div class="modal-box">
            <div class="modal-header">
                <h3 id="qvTitle">تفاصيل المنحة والمطابقة</h3>
                <button class="close-modal-btn" onclick="closeQuickView()">&times;</button>
            </div>
            <div id="qvContent"></div>
        </div>
    </div>

    <!-- Toast Notification -->
    <div class="toast" id="toastBox">
        <span id="toastIcon">✅</span>
        <span id="toastMsg">تمت العملية بنجاح!</span>
    </div>

    <!-- Application Script -->
    <script>
        const INITIAL_DATA = """ + data_json_str + """;
        const STORAGE_KEY = 'radwa_scholarships_data_v4';
        
        let scholarships = [];
        let activeCategory = 'all';

        // Load data from localStorage or fallback to INITIAL_DATA
        function loadData() {
            try {
                const saved = localStorage.getItem(STORAGE_KEY);
                if (saved) {
                    scholarships = JSON.parse(saved);
                } else {
                    scholarships = JSON.parse(JSON.stringify(INITIAL_DATA));
                }
            } catch (e) {
                console.error("Storage load error:", e);
                scholarships = JSON.parse(JSON.stringify(INITIAL_DATA));
            }
        }

        // Save data to localStorage
        function saveToStorage(notify = false) {
            try {
                localStorage.setItem(STORAGE_KEY, JSON.stringify(scholarships));
                if (notify) {
                    showToast('تم حفظ جميع تعديلاتك بنجاح في المتصفح!', '💾');
                    document.getElementById('editNoticeBar').classList.remove('active');
                }
            } catch (e) {
                console.error("Storage save error:", e);
            }
        }

        // Reset to original data
        function resetToOriginal() {
            if (confirm('هل أنتِ متأكدة من رغبتك في استعادة قاعدة البيانات الأصلية والتراجع عن أي تعديلات قمتِ بها؟')) {
                localStorage.removeItem(STORAGE_KEY);
                scholarships = JSON.parse(JSON.stringify(INITIAL_DATA));
                renderTable();
                updateStats();
                showToast('تمت استعادة قاعدة البيانات الأصلية بنجاح!', '🔄');
            }
        }

        // Category Badge Helper
        function getCategoryBadge(cat) {
            if (cat.includes('أفريقيا') || cat.includes('النامية')) return '<span class="badge badge-africa">🌍 أفريقيا والدول النامية</span>';
            if (cat.includes('حكومية') || cat.includes('دولية')) return '<span class="badge badge-gov">🏛️ منحة حكومية عالمية</span>';
            if (cat.includes('تكنولوجيا') || cat.includes('العلوم') || cat.includes('بيوتكنولوجي')) return '<span class="badge badge-biotech">🧬 بيوتكنولوجي وأبحاث</span>';
            if (cat.includes('جامعات') || cat.includes('مؤسسات')) return '<span class="badge badge-uni">🎓 جامعة كبرى</span>';
            return '<span class="badge badge-dir">💼 وظيفة / دليل بحث</span>';
        }

        // Match Rate Badge Helper
        function getMatchBadge(rate, status) {
            const num = parseInt(rate) || 90;
            let cls = 'match-90';
            if (num >= 95) cls = 'match-95';
            else if (num >= 90) cls = 'match-90';
            else if (num >= 85) cls = 'match-85';
            else cls = 'match-80';

            return `
                <div style="text-align: center;">
                    <div class="match-badge ${cls}">🎯 توافق ${rate}</div>
                    <span class="readiness-text">${status || 'مؤهلة للتقديم'}</span>
                </div>
            `;
        }

        // Render Table Rows
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const searchVal = document.getElementById('searchInput').value.toLowerCase().trim();
            const degreeVal = document.getElementById('degreeFilter').value;
            const matchVal = document.getElementById('matchFilter').value;
            const fundingVal = document.getElementById('fundingFilter').value;

            tbody.innerHTML = '';

            let count = 0;

            scholarships.forEach((item, index) => {
                // Category Filter
                if (activeCategory !== 'all' && item.category !== activeCategory) {
                    return;
                }

                // Degree Filter
                if (degreeVal !== 'all' && !item.degree.includes(degreeVal)) {
                    return;
                }

                // Match Rate Filter
                if (matchVal !== 'all') {
                    const r = parseInt(item.match_rate) || 90;
                    if (r < parseInt(matchVal)) return;
                }

                // Funding Filter
                if (fundingVal !== 'all') {
                    if (fundingVal === 'تمويل كامل' && !item.funding.includes('كامل')) return;
                    if (fundingVal === 'وظيفة' && !item.funding.includes('وظيفة') && !item.degree.includes('وظيفة')) return;
                }

                // Search Filter
                if (searchVal) {
                    const textSearch = (item.name + ' ' + (item.name_ar||'') + ' ' + item.category + ' ' + (item.detailed_requirements||'') + ' ' + (item.cv_gap_analysis||'') + ' ' + item.start_date + ' ' + item.deadline + ' ' + item.funding + ' ' + (item.notes||'')).toLowerCase();
                    if (!textSearch.includes(searchVal)) {
                        return;
                    }
                }

                count++;

                const tr = document.createElement('tr');
                tr.setAttribute('data-id', item.id);

                const displayName = item.name_ar && item.name_ar !== item.name ? 
                    `<div class="scholarship-title">${item.name_ar}</div><div class="scholarship-title-en">${item.name}</div>` : 
                    `<div class="scholarship-title">${item.name}</div>`;

                const reqText = item.detailed_requirements || item.requirements || 'راجع الموقع الرسمي';
                const gapText = item.cv_gap_analysis || 'تطابق مباشر مع ملفكِ الأكاديمي';

                tr.innerHTML = `
                    <td style="text-align: center; color: var(--text-muted); font-weight: 700;">${count}</td>
                    <td>
                        ${displayName}
                        <div class="tags-row">
                            ${getCategoryBadge(item.category)}
                            <span class="badge badge-degree">${item.degree}</span>
                        </div>
                        <div style="margin-top: 8px;">
                            <a href="${item.link}" target="_blank" rel="noopener noreferrer" class="link-pill" title="فتح الموقع الرسمي">
                                🔗 الموقع الرسمي
                            </a>
                        </div>
                    </td>
                    <td>
                        ${getMatchBadge(item.match_rate, item.readiness_status)}
                    </td>
                    <td>
                        <div class="req-box" contenteditable="true" oninput="markEdited()">${reqText}</div>
                    </td>
                    <td>
                        <div class="gap-box" contenteditable="true" oninput="markEdited()">${gapText}</div>
                    </td>
                    <td>
                        <span class="date-box">📅 <strong>البداية:</strong> ${item.start_date}</span>
                        <span class="date-box date-deadline">⏰ <strong>آخر موعد:</strong> ${item.deadline}</span>
                        <div style="font-size: 11px; color: var(--text-muted); margin-top: 6px;">
                            💰 ${item.funding}
                        </div>
                    </td>
                    <td style="text-align: center;">
                        <div class="action-btns-cell">
                            <button class="icon-btn" onclick="openQuickView('${item.id}')" title="عرض بطاقة التفاصيل الشاملة">👁️</button>
                            <button class="icon-btn" onclick="copyLink('${item.link}')" title="نسخ رابط المنحة">📋</button>
                            <button class="icon-btn delete-btn" onclick="deleteItem('${item.id}')" title="حذف المنحة">🗑️</button>
                        </div>
                    </td>
                `;

                tbody.appendChild(tr);
            });

            if (count === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="7" style="text-align: center; padding: 40px; color: var(--text-muted);">
                            <div style="font-size: 32px; margin-bottom: 8px;">🔍</div>
                            <strong>لم يتم العثور على أي منحة تطابق البحث أو الفلاتر المختارة!</strong>
                            <p style="font-size: 12px; margin-top: 4px;">جربي إزالة بعض الكلمات من شريط البحث أو تغيير خيارات التوافق والتصنيف.</p>
                        </td>
                    </tr>
                `;
            }
        }

        // Quick View Modal
        function openQuickView(id) {
            const item = scholarships.find(s => s.id === id);
            if (!item) return;

            document.getElementById('qvTitle').innerText = item.name_ar || item.name;
            const content = `
                <div style="margin-bottom: 16px;">
                    <div style="display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap;">
                        ${getCategoryBadge(item.category)}
                        <span class="badge badge-degree">${item.degree}</span>
                        <span class="match-badge match-95">🎯 نسبة التوافق: ${item.match_rate || '90%'}</span>
                    </div>
                    <p style="font-size: 13px; color: var(--text-muted); font-family: 'Plus Jakarta Sans', sans-serif;">${item.name}</p>
                </div>

                <div style="margin-bottom: 16px;">
                    <h4 style="font-size: 14px; font-weight: 800; color: var(--accent); margin-bottom: 6px;">📋 المتطلبات التفصيلية والشروط:</h4>
                    <div class="req-box" style="max-width: 100%; font-size: 13px;">${item.detailed_requirements || item.requirements}</div>
                </div>

                <div style="margin-bottom: 16px;">
                    <h4 style="font-size: 14px; font-weight: 800; color: var(--primary); margin-bottom: 6px;">⚖️ طرح الـ CV الخاص بكِ والمتبقي لاستيفاء الشروط 100%:</h4>
                    <div class="gap-box" style="max-width: 100%; font-size: 13px;">${item.cv_gap_analysis}</div>
                </div>

                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; background: var(--bg-primary); padding: 12px; border-radius: var(--radius-sm); margin-bottom: 20px;">
                    <div>📅 <strong>ميعاد البداية:</strong> ${item.start_date}</div>
                    <div style="color: #dc2626;">⏰ <strong>الموعد النهائي:</strong> ${item.deadline}</div>
                    <div style="grid-column: span 2;">💰 <strong>نوع التمويل:</strong> ${item.funding}</div>
                </div>

                <div style="display: flex; justify-content: flex-end; gap: 10px;">
                    <a href="${item.link}" target="_blank" class="btn btn-primary">🔗 الانتقال للموقع الرسمي</a>
                    <button class="btn btn-secondary" onclick="closeQuickView()">إغلاق</button>
                </div>
            `;
            document.getElementById('qvContent').innerHTML = content;
            document.getElementById('quickViewModal').classList.add('active');
        }

        function closeQuickView() {
            document.getElementById('quickViewModal').classList.remove('active');
        }

        // Track when user edits a cell
        function markEdited() {
            document.getElementById('editNoticeBar').classList.add('active');
        }

        // Update statistics
        function updateStats() {
            const total = scholarships.length;
            const match95 = scholarships.filter(s => (parseInt(s.match_rate) || 0) >= 95).length;
            const africa = scholarships.filter(s => s.category.includes('أفريقيا') || s.category.includes('النامية')).length;
            const biotech = scholarships.filter(s => s.category.includes('تكنولوجيا') || (s.detailed_requirements||'').toLowerCase().includes('biotech')).length;
            const phd = scholarships.filter(s => s.degree.includes('دكتوراه') || s.degree.includes('وظيفة') || s.category.includes('محركات')).length;

            document.getElementById('statTotal').innerText = total;
            document.getElementById('statMatch95').innerText = match95;
            document.getElementById('statAfrica').innerText = africa;
            document.getElementById('statBiotech').innerText = biotech;
            document.getElementById('statPhd').innerText = phd;

            // Tab counts
            document.getElementById('countAll').innerText = total;
            document.getElementById('countAfrica').innerText = scholarships.filter(s => s.category === 'أفريقيا والدول النامية').length;
            document.getElementById('countGov').innerText = scholarships.filter(s => s.category === 'المنح الحكومية والدولية الكبرى').length;
            document.getElementById('countBiotechTab').innerText = scholarships.filter(s => s.category === 'مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية').length;
            document.getElementById('countUni').innerText = scholarships.filter(s => s.category === 'منح الجامعات العالمية والمؤسسات المرموقة').length;
            document.getElementById('countDir').innerText = scholarships.filter(s => s.category === 'محركات البحث وأدلة وظائف الدكتوراه').length;
        }

        // Set category filter
        function setCategory(cat, btn) {
            activeCategory = cat;
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderTable();
        }

        // Apply filters
        function applyFilters() {
            renderTable();
        }

        // Live search
        function handleSearch() {
            renderTable();
        }

        // Copy link to clipboard
        function copyLink(url) {
            if (!url) return;
            navigator.clipboard.writeText(url).then(() => {
                showToast('تم نسخ الرابط بنجاح إلى الحافظة!', '📋');
            }).catch(() => {
                const el = document.createElement('textarea');
                el.value = url;
                document.body.appendChild(el);
                el.select();
                document.execCommand('copy');
                document.body.removeChild(el);
                showToast('تم نسخ الرابط بنجاح!', '📋');
            });
        }

        // Delete item
        function deleteItem(id) {
            if (confirm('هل أنتِ متأكدة من رغبتك في حذف هذه الفرصة من القائمة؟')) {
                scholarships = scholarships.filter(s => s.id !== id);
                renderTable();
                updateStats();
                saveToStorage();
                showToast('تم حذف المنحة من القائمة بنجاح!', '🗑️');
            }
        }

        // Modal Controls
        function openAddModal() {
            document.getElementById('addModal').classList.add('active');
        }

        function closeAddModal() {
            document.getElementById('addModal').classList.remove('active');
        }

        // Handle Add New
        function handleAddNew(e) {
            e.preventDefault();
            const name = document.getElementById('newName').value.trim();
            const category = document.getElementById('newCategory').value;
            const degree = document.getElementById('newDegree').value;
            const link = document.getElementById('newLink').value.trim();
            const reqs = document.getElementById('newReqs').value.trim();
            const gap = document.getElementById('newGap').value.trim() || 'تطابق مباشر مع ملفكِ الأكاديمي';
            const start = document.getElementById('newStart').value.trim() || 'حسب دورة البرنامج السنوية';
            const deadline = document.getElementById('newDeadline').value.trim() || 'راجع الموقع الرسمي';
            const notes = document.getElementById('newNotes').value.trim() || 'تمويل كامل أو جزئي';

            const newEntry = {
                id: 'CUSTOM-' + Date.now(),
                name: name,
                name_ar: name,
                category: category,
                degree: degree,
                link: link,
                detailed_requirements: reqs,
                cv_gap_analysis: gap,
                match_rate: '90%',
                readiness_status: 'مؤهلة للتقديم',
                start_date: start,
                deadline: deadline,
                funding: notes,
                notes: ''
            };

            scholarships.unshift(newEntry);
            saveToStorage();
            renderTable();
            updateStats();
            closeAddModal();
            showToast('تمت إضافة المنحة بنجاح إلى قاعدة البيانات!', '🎉');
            document.getElementById('addForm').reset();
        }

        // Theme Toggle
        function toggleTheme() {
            const current = document.body.getAttribute('data-theme');
            const next = current === 'dark' ? 'light' : 'dark';
            document.body.setAttribute('data-theme', next);
            localStorage.setItem('radwa_scholarships_theme', next);
            document.getElementById('themeIcon').innerText = next === 'dark' ? '☀️' : '🌙';
        }

        // Toast Message
        function showToast(msg, icon = '✅') {
            const toast = document.getElementById('toastBox');
            document.getElementById('toastMsg').innerText = msg;
            document.getElementById('toastIcon').innerText = icon;
            toast.classList.add('show');
            setTimeout(() => {
                toast.classList.remove('show');
            }, 3000);
        }

        // Export directly to Excel (.xlsx) using SheetJS
        function exportToExcel() {
            saveToStorage();
            const exportData = scholarships.map((item, idx) => ({
                "#": idx + 1,
                "اسم المنحة / المنظمة": item.name_ar || item.name,
                "التصنيف": item.category,
                "المرحلة الدراسية": item.degree,
                "نسبة التوافق مع الـ CV": item.match_rate || "90%",
                "حالة الاستيفاء والجاهزية": item.readiness_status || "مؤهلة للتقديم",
                "المتطلبات التفصيلية والشروط الأكاديمية": item.detailed_requirements || item.requirements,
                "طرح الـ CV: المتبقي لرضوى لاستيفاء الشروط 100%": item.cv_gap_analysis,
                "الرابط الرسمي": item.link,
                "ميعاد بداية التقديم": item.start_date,
                "آخر موعد للتقديم": item.deadline,
                "نوع التمويل والملاحظات": item.funding + (item.notes ? " | " + item.notes : "")
            }));

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "المنح والمطابقة لرضوى");
            XLSX.writeFile(wb, "Scholarships_Master_PhD_Radwa_CV_Matched.xlsx");
            showToast('تم تصدير ملف الإكسيل مع مطابقة الـ CV بنجاح!', '📊');
        }

        // Init on load
        window.addEventListener('DOMContentLoaded', () => {
            const savedTheme = localStorage.getItem('radwa_scholarships_theme') || 'light';
            document.body.setAttribute('data-theme', savedTheme);
            document.getElementById('themeIcon').innerText = savedTheme === 'dark' ? '☀️' : '🌙';
            
            loadData();
            updateStats();
            renderTable();
        });
    </script>
</body>
</html>
"""

# Write to both target filenames:
target_html_1 = r"c:\Users\radwa kamal\Downloads\Masters & PHD Radwa\Scholarships_Master_PhD_Radwa_Interactive.html"
target_html_2 = r"c:\Users\radwa kamal\Downloads\Masters & PHD Radwa\index.html"

with open(target_html_1, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Written: {target_html_1}")

with open(target_html_2, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Written: {target_html_2}")
