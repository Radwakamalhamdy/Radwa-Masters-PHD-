# -*- coding: utf-8 -*-
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
    <title>بوابة وقاعدة بيانات منح الماجستير والدكتوراه | رضوى كمال</title>
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
    
    <!-- SheetJS for Instant Excel Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <style>
        :root {
            --bg-primary: #f8fafc;
            --bg-secondary: #ffffff;
            --bg-card: rgba(255, 255, 255, 0.95);
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
            --bg-card: rgba(17, 24, 39, 0.75);
            --bg-card-hover: rgba(31, 41, 55, 0.85);
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
            background: radial-gradient(circle, rgba(5, 150, 105, 0.08) 0%, rgba(37, 99, 235, 0.04) 60%, transparent 70%);
            filter: blur(80px);
            z-index: -1;
            pointer-events: none;
        }

        .ambient-glow-2 {
            position: fixed;
            bottom: -200px;
            left: 10%;
            width: 500px;
            height: 500px;
            background: radial-gradient(circle, rgba(124, 58, 237, 0.06) 0%, rgba(217, 119, 6, 0.04) 60%, transparent 70%);
            filter: blur(80px);
            z-index: -1;
            pointer-events: none;
        }

        .container {
            max-width: 1600px;
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
            width: 54px;
            height: 54px;
            border-radius: 14px;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            box-shadow: 0 8px 16px rgba(5, 150, 105, 0.25);
            color: #fff;
        }

        .header-title h1 {
            font-size: 23px;
            font-weight: 800;
            color: var(--text-main);
            margin-bottom: 4px;
        }

        .header-title p {
            font-size: 13px;
            color: var(--text-muted);
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 12px;
            flex-wrap: wrap;
        }

        /* Buttons */
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 10px 18px;
            font-size: 13px;
            font-weight: 700;
            border-radius: var(--radius-sm);
            border: none;
            cursor: pointer;
            transition: var(--transition);
            text-decoration: none;
        }

        .btn-primary {
            background: linear-gradient(135deg, var(--primary), var(--primary-hover));
            color: white;
            box-shadow: 0 4px 12px rgba(5, 150, 105, 0.2);
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(5, 150, 105, 0.35);
        }

        .btn-secondary {
            background: var(--bg-secondary);
            color: var(--text-main);
            border: 1px solid var(--border-color);
        }

        .btn-secondary:hover {
            background: var(--bg-card-hover);
            border-color: var(--primary);
            transform: translateY(-2px);
        }

        .theme-toggle-btn {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            cursor: pointer;
            font-size: 18px;
            transition: var(--transition);
        }

        .theme-toggle-btn:hover {
            border-color: var(--warning);
            transform: rotate(20deg);
        }

        /* Stats KPI Bar */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .stat-card {
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 18px 20px;
            display: flex;
            align-items: center;
            gap: 16px;
            box-shadow: var(--shadow);
            transition: var(--transition);
        }

        .stat-card:hover {
            transform: translateY(-3px);
            border-color: rgba(5, 150, 105, 0.3);
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

        .stat-icon.total { background: var(--primary-light); color: var(--primary); }
        .stat-icon.africa { background: var(--warning-light); color: var(--warning); }
        .stat-icon.masters { background: var(--accent-light); color: var(--accent); }
        .stat-icon.phd { background: var(--purple-light); color: var(--purple); }
        .stat-icon.biotech { background: var(--danger-light); color: var(--danger); }

        .stat-info h3 {
            font-size: 22px;
            font-weight: 800;
            line-height: 1.1;
        }

        .stat-info p {
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 3px;
        }

        /* Controls & Filters Area */
        .controls-panel {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            padding: 20px 24px;
            margin-bottom: 24px;
            box-shadow: var(--shadow);
            display: flex;
            flex-direction: column;
            gap: 16px;
        }

        .search-and-select {
            display: flex;
            gap: 14px;
            flex-wrap: wrap;
        }

        .search-wrapper {
            flex: 1;
            min-width: 280px;
            position: relative;
        }

        .search-icon {
            position: absolute;
            right: 16px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            font-size: 16px;
            pointer-events: none;
        }

        .search-input {
            width: 100%;
            padding: 12px 46px 12px 16px;
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-sm);
            color: var(--text-main);
            font-size: 14px;
            outline: none;
            transition: var(--transition);
        }

        .search-input:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(5, 150, 105, 0.15);
        }

        .select-filter {
            padding: 12px 18px;
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-sm);
            color: var(--text-main);
            font-size: 14px;
            outline: none;
            cursor: pointer;
            min-width: 180px;
        }

        .select-filter:focus {
            border-color: var(--accent);
        }

        /* Category Filter Tabs */
        .category-tabs {
            display: flex;
            gap: 8px;
            overflow-x: auto;
            padding-bottom: 4px;
            scrollbar-width: thin;
        }

        .tab-btn {
            background: var(--bg-secondary);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 8px 16px;
            border-radius: 30px;
            font-size: 12.5px;
            font-weight: 600;
            cursor: pointer;
            white-space: nowrap;
            transition: var(--transition);
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .tab-btn:hover {
            color: var(--text-main);
            border-color: var(--primary);
        }

        .tab-btn.active {
            background: var(--primary);
            color: white;
            border-color: var(--primary);
            box-shadow: 0 4px 12px rgba(5, 150, 105, 0.25);
        }

        /* Edit status banner */
        .edit-notice-bar {
            display: none;
            background: linear-gradient(90deg, rgba(217, 119, 6, 0.12), rgba(5, 150, 105, 0.12));
            border: 1px solid var(--warning);
            border-radius: var(--radius-sm);
            padding: 12px 20px;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            animation: fadeIn 0.3s ease;
        }

        .edit-notice-bar.active {
            display: flex;
        }

        /* Table Container */
        .table-responsive {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow);
            overflow-x: auto;
            position: relative;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: right;
            font-size: 13px;
        }

        th {
            background: #f1f5f9;
            color: #334155;
            font-weight: 700;
            padding: 14px 16px;
            border-bottom: 2px solid var(--border-color);
            position: sticky;
            top: 0;
            z-index: 10;
            white-space: nowrap;
        }

        [data-theme="dark"] th {
            background: rgba(15, 23, 42, 0.95);
            color: #94a3b8;
        }

        td {
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            vertical-align: middle;
            transition: background 0.15s ease;
            background: var(--bg-secondary);
        }

        tr:nth-child(even) td {
            background: rgba(248, 250, 252, 0.7);
        }

        [data-theme="dark"] tr:nth-child(even) td {
            background: rgba(30, 41, 59, 0.4);
        }

        tr:hover td {
            background: var(--bg-card-hover) !important;
        }

        /* Editable cell styles */
        td[contenteditable="true"] {
            outline: none;
            border-radius: 4px;
            transition: all 0.2s;
        }

        td[contenteditable="true"]:focus {
            background: rgba(5, 150, 105, 0.08) !important;
            box-shadow: inset 0 0 0 1.5px var(--primary);
        }

        /* Badge Pills */
        .badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 700;
            white-space: nowrap;
        }

        .badge-africa { background: var(--warning-light); color: var(--warning); border: 1px solid rgba(217, 119, 6, 0.25); }
        .badge-gov { background: var(--accent-light); color: var(--accent); border: 1px solid rgba(37, 99, 235, 0.25); }
        .badge-biotech { background: var(--primary-light); color: var(--primary); border: 1px solid rgba(5, 150, 105, 0.25); }
        .badge-uni { background: var(--purple-light); color: var(--purple); border: 1px solid rgba(124, 58, 237, 0.25); }
        .badge-dir { background: rgba(236, 72, 153, 0.12); color: #db2777; border: 1px solid rgba(236, 72, 153, 0.25); }

        .degree-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            color: var(--text-main);
        }

        .link-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: var(--accent);
            text-decoration: none;
            font-weight: 700;
            font-size: 12px;
            padding: 5px 10px;
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
            font-size: 12px;
            font-weight: 600;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 4px;
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

        /* Modal Dialog */
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.6);
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
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            width: 100%;
            max-width: 650px;
            padding: 28px;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15);
            transform: translateY(20px);
            transition: all 0.3s ease;
            max-height: 90vh;
            overflow-y: auto;
        }

        .modal-overlay.active .modal-box {
            transform: translateY(0);
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 1px solid var(--border-color);
        }

        .modal-header h3 {
            font-size: 18px;
            font-weight: 800;
        }

        .close-modal-btn {
            background: none;
            border: none;
            font-size: 22px;
            color: var(--text-muted);
            cursor: pointer;
        }

        .close-modal-btn:hover {
            color: var(--danger);
        }

        .form-group {
            margin-bottom: 14px;
        }

        .form-group label {
            display: block;
            font-size: 12.5px;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 6px;
        }

        .form-control {
            width: 100%;
            padding: 10px 14px;
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-sm);
            color: var(--text-main);
            font-size: 13.5px;
            outline: none;
        }

        .form-control:focus {
            border-color: var(--primary);
        }

        textarea.form-control {
            min-height: 80px;
            resize: vertical;
        }

        /* Toast notification */
        .toast {
            position: fixed;
            bottom: 30px;
            left: 30px;
            background: #0f172a;
            color: white;
            border-right: 4px solid var(--primary);
            border-radius: var(--radius-sm);
            padding: 14px 20px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 13px;
            font-weight: 600;
            transform: translateY(100px);
            opacity: 0;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            z-index: 1000;
        }

        .toast.show {
            transform: translateY(0);
            opacity: 1;
        }

        /* Responsive tweaks */
        @media (max-width: 768px) {
            header {
                flex-direction: column;
                align-items: flex-start;
            }
            .header-actions {
                width: 100%;
                justify-content: space-between;
            }
            .btn {
                flex: 1;
                justify-content: center;
            }
        }

        /* Print styles */
        @media print {
            body { background: white !important; color: black !important; }
            header, .controls-panel, .ambient-glow, .ambient-glow-2, .action-btns-cell, .theme-toggle-btn { display: none !important; }
            .table-responsive { border: none !important; box-shadow: none !important; }
            th, td { border: 1px solid #ccc !important; color: black !important; }
        }
    </style>
</head>
<body data-theme="light">

    <div class="ambient-glow"></div>
    <div class="ambient-glow-2"></div>

    <div class="container">
        <!-- Top Header -->
        <header>
            <div class="header-title-box">
                <div class="header-badge-icon">🎓</div>
                <div class="header-title">
                    <h1>بوابة وقاعدة بيانات منح الماجستير والدكتوراه</h1>
                    <p>موسوعة شاملة للمنظمات العالمية، الأفريقية، والدول النامية — إعداد مخصص للباحثة / رضوى كمال</p>
                </div>
            </div>
            <div class="header-actions">
                <button class="btn btn-primary" onclick="openAddModal()">
                    <span>➕</span> إضافة منحة جديدة
                </button>
                <button class="btn btn-secondary" onclick="exportToExcel()">
                    <span>📊</span> تصدير إلى Excel
                </button>
                <button class="btn btn-secondary" onclick="window.print()">
                    <span>🖨️</span> طباعة / PDF
                </button>
                <button class="btn btn-secondary" id="resetBtn" onclick="resetToOriginal()">
                    <span>🔄</span> استعادة الأصل
                </button>
                <button class="theme-toggle-btn" onclick="toggleTheme()" title="تبديل المظهر النهاري/الليلي">
                    <span id="themeIcon">🌙</span>
                </button>
            </div>
        </header>

        <!-- Stats KPI Bar -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon total">🏛️</div>
                <div class="stat-info">
                    <h3 id="statTotal">0</h3>
                    <p>إجمالي الفرص والمنظمات</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon africa">🌍</div>
                <div class="stat-info">
                    <h3 id="statAfrica">0</h3>
                    <p>أفريقيا والدول النامية</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon masters">📚</div>
                <div class="stat-info">
                    <h3 id="statMasters">0</h3>
                    <p>برامج الماجستير</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon phd">🔬</div>
                <div class="stat-info">
                    <h3 id="statPhd">0</h3>
                    <p>الدكتوراه والزمالات البحثية</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon biotech">🧬</div>
                <div class="stat-info">
                    <h3 id="statBiotech">0</h3>
                    <p>مراكز البيوتكنولوجي والعلوم</p>
                </div>
            </div>
        </div>

        <!-- Controls & Filters Panel -->
        <div class="controls-panel">
            <div class="search-and-select">
                <div class="search-wrapper">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="searchInput" class="search-input" placeholder="ابحثي بأي كلمة: اسم المنحة، الدولة، الشروط، الرابط، المواعيد..." oninput="handleSearch()">
                </div>
                <select id="degreeFilter" class="select-filter" onchange="applyFilters()">
                    <option value="all">كل المراحل الدراسية</option>
                    <option value="ماجستير">ماجستير فقط</option>
                    <option value="دكتوراه">دكتوراه فقط</option>
                    <option value="ماجستير ودكتوراه">ماجستير ودكتوراه معاً</option>
                </select>
                <select id="fundingFilter" class="select-filter" onchange="applyFilters()">
                    <option value="all">كل أنواع التمويل</option>
                    <option value="تمويل كامل">تمويل كامل (100%)</option>
                    <option value="وظيفة">وظيفة بحثية براتب</option>
                </select>
            </div>

            <!-- Category Tabs -->
            <div class="category-tabs" id="categoryTabs">
                <button class="tab-btn active" onclick="setCategory('all', this)">⭐ عرض الكل (<span id="countAll">0</span>)</button>
                <button class="tab-btn" onclick="setCategory('أفريقيا والدول النامية', this)">🌍 أفريقيا والدول النامية (<span id="countAfrica">0</span>)</button>
                <button class="tab-btn" onclick="setCategory('المنح الحكومية والدولية الكبرى', this)">🏛️ المنح الحكومية العالمية (<span id="countGov">0</span>)</button>
                <button class="tab-btn" onclick="setCategory('مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية', this)">🧬 البيوتكنولوجي والمراكز البحثية (<span id="countBiotechTab">0</span>)</button>
                <button class="tab-btn" onclick="setCategory('منح الجامعات العالمية والمؤسسات المرموقة', this)">🎓 الجامعات والمؤسسات الكبرى (<span id="countUni">0</span>)</button>
                <button class="tab-btn" onclick="setCategory('محركات البحث وأدلة وظائف الدكتوراه', this)">🔍 محركات وأدلة الدكتوراه (<span id="countDir">0</span>)</button>
            </div>
        </div>

        <!-- Edit notice & Auto-Save bar -->
        <div class="edit-notice-bar" id="editNoticeBar">
            <div style="display: flex; align-items: center; gap: 8px;">
                <span>✏️</span>
                <span><strong>وضع التعديل المباشر نشط:</strong> يمكنك الضغط على أي خلية في الجدول لتعديلها. التعديلات تحفظ تلقائياً في متصفحك!</span>
            </div>
            <button class="btn btn-primary" onclick="saveToStorage(true)">
                <span>💾</span> حفظ التعديلات الآن
            </button>
        </div>

        <!-- Main Table -->
        <div class="table-responsive">
            <table id="scholarshipsTable">
                <thead>
                    <tr>
                        <th style="width: 45px; text-align: center;">#</th>
                        <th style="width: 260px;">المنحة / المنظمة</th>
                        <th style="width: 170px;">التصنيف</th>
                        <th style="width: 130px;">المرحلة</th>
                        <th style="width: 140px;">الرابط الرسمي</th>
                        <th style="min-width: 300px;">الشروط الأساسية والمتطلبات</th>
                        <th style="width: 160px;">ميعاد البداية</th>
                        <th style="width: 160px;">آخر موعد (Deadline)</th>
                        <th style="min-width: 220px;">التمويل والملاحظات</th>
                        <th style="width: 90px; text-align: center;">إجراءات</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <!-- Populated via Javascript -->
                </tbody>
            </table>
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
                        <option value="زمالة / وظيفة بحثية">زمالة / وظيفة بحثية</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>الرابط الرسمي *</label>
                    <input type="url" id="newLink" class="form-control" placeholder="https://..." required>
                </div>
                <div class="form-group">
                    <label>الشروط الأساسية والمتطلبات *</label>
                    <textarea id="newReqs" class="form-control" placeholder="اكتبي المؤهل المطلوب، شرط السن، اللغة، التخصصات..." required></textarea>
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
                    <label>نوع التمويل والملاحظات</label>
                    <input type="text" id="newNotes" class="form-control" placeholder="مثال: تمويل كامل 100% يشمل السكن والراتب">
                </div>
                <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px;">
                    <button type="button" class="btn btn-secondary" onclick="closeAddModal()">إلغاء</button>
                    <button type="submit" class="btn btn-primary">حفظ الإضافة</button>
                </div>
            </form>
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
        const STORAGE_KEY = 'radwa_scholarships_data_v3';
        
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
                const rows = document.querySelectorAll('#tableBody tr');
                if (rows.length > 0) {
                    rows.forEach(row => {
                        const id = row.getAttribute('data-id');
                        const item = scholarships.find(s => s.id === id);
                        if (item) {
                            item.name_ar = row.querySelector('.col-name').innerText.trim();
                            item.requirements = row.querySelector('.col-reqs').innerText.trim();
                            item.start_date = row.querySelector('.col-start').innerText.trim();
                            item.deadline = row.querySelector('.col-deadline').innerText.trim();
                            item.funding = row.querySelector('.col-notes').innerText.trim();
                        }
                    });
                }
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
            if (cat.includes('جامعات') || cat.includes('مؤسسات')) return '<span class="badge badge-uni">🎓 جامعة / مؤسسة كبرى</span>';
            return '<span class="badge badge-dir">🔍 محرك / دليل بحث</span>';
        }

        // Render Table Rows
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const searchVal = document.getElementById('searchInput').value.toLowerCase().trim();
            const degreeVal = document.getElementById('degreeFilter').value;
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

                // Funding Filter
                if (fundingVal !== 'all') {
                    if (fundingVal === 'تمويل كامل' && !item.funding.includes('كامل')) return;
                    if (fundingVal === 'وظيفة' && !item.funding.includes('وظيفة') && !item.degree.includes('وظيفة')) return;
                }

                // Search Filter
                if (searchVal) {
                    const textSearch = (item.name + ' ' + (item.name_ar||'') + ' ' + item.category + ' ' + item.requirements + ' ' + item.start_date + ' ' + item.deadline + ' ' + item.funding + ' ' + (item.notes||'')).toLowerCase();
                    if (!textSearch.includes(searchVal)) {
                        return;
                    }
                }

                count++;

                const tr = document.createElement('tr');
                tr.setAttribute('data-id', item.id);

                const displayName = item.name_ar && item.name_ar !== item.name ? `<strong>${item.name_ar}</strong><br><small style="color:var(--text-muted); font-family: 'Plus Jakarta Sans', sans-serif;">${item.name}</small>` : `<strong>${item.name}</strong>`;

                tr.innerHTML = `
                    <td style="text-align: center; color: var(--text-muted); font-weight: 700;">${count}</td>
                    <td class="col-name" contenteditable="true" oninput="markEdited()">${displayName}</td>
                    <td>${getCategoryBadge(item.category)}</td>
                    <td><span class="degree-badge">${item.degree}</span></td>
                    <td>
                        <a href="${item.link}" target="_blank" rel="noopener noreferrer" class="link-pill" title="فتح الموقع الرسمي">
                            🔗 فتح الموقع
                        </a>
                    </td>
                    <td class="col-reqs" contenteditable="true" oninput="markEdited()" style="line-height: 1.5; font-size: 12.5px;">${item.requirements}</td>
                    <td class="col-start" contenteditable="true" oninput="markEdited()"><span class="date-box">📅 ${item.start_date}</span></td>
                    <td class="col-deadline" contenteditable="true" oninput="markEdited()"><span class="date-box date-deadline">⏰ ${item.deadline}</span></td>
                    <td class="col-notes" contenteditable="true" oninput="markEdited()" style="font-size: 12px; color: var(--text-muted);">${item.funding} ${item.notes ? '<br><small>💡 ' + item.notes + '</small>' : ''}</td>
                    <td style="text-align: center;">
                        <div class="action-btns-cell">
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
                        <td colspan="10" style="text-align: center; padding: 40px; color: var(--text-muted);">
                            <div style="font-size: 32px; margin-bottom: 8px;">🔍</div>
                            <strong>لم يتم العثور على أي منحة تطابق البحث أو الفلاتر المختارة!</strong>
                            <p style="font-size: 12px; margin-top: 4px;">جربي إزالة بعض الكلمات من شريط البحث أو تغيير التصنيف.</p>
                        </td>
                    </tr>
                `;
            }
        }

        // Track when user edits a cell
        function markEdited() {
            document.getElementById('editNoticeBar').classList.add('active');
        }

        // Update statistics
        function updateStats() {
            const total = scholarships.length;
            const africa = scholarships.filter(s => s.category.includes('أفريقيا') || s.category.includes('النامية')).length;
            const masters = scholarships.filter(s => s.degree.includes('ماجستير')).length;
            const phd = scholarships.filter(s => s.degree.includes('دكتوراه') || s.degree.includes('وظيفة')).length;
            const biotech = scholarships.filter(s => s.category.includes('تكنولوجيا') || s.requirements.toLowerCase().includes('biotech') || s.name.toLowerCase().includes('biotech')).length;

            document.getElementById('statTotal').innerText = total;
            document.getElementById('statAfrica').innerText = africa;
            document.getElementById('statMasters').innerText = masters;
            document.getElementById('statPhd').innerText = phd;
            document.getElementById('statBiotech').innerText = biotech;

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

        // Delete an item
        function deleteItem(id) {
            if (confirm('هل أنتِ متأكدة من حذف هذه الفرصة؟')) {
                scholarships = scholarships.filter(s => s.id !== id);
                saveToStorage();
                renderTable();
                updateStats();
                showToast('تم حذف المنحة بنجاح!', '🗑️');
            }
        }

        // Modal Controls
        function openAddModal() {
            document.getElementById('addModal').classList.add('active');
        }

        function closeAddModal() {
            document.getElementById('addModal').classList.remove('active');
        }

        // Handle Add New Form Submit
        function handleAddNew(e) {
            e.preventDefault();
            const newObj = {
                id: 'CUSTOM-' + Date.now(),
                name: document.getElementById('newName').value.trim(),
                name_ar: document.getElementById('newName').value.trim(),
                category: document.getElementById('newCategory').value,
                degree: document.getElementById('newDegree').value,
                link: document.getElementById('newLink').value.trim(),
                requirements: document.getElementById('newReqs').value.trim(),
                start_date: document.getElementById('newStart').value.trim() || 'حسب الدورة السنوية',
                deadline: document.getElementById('newDeadline').value.trim() || 'راجع الموقع الرسمي',
                funding: document.getElementById('newNotes').value.trim() || 'تمويل دراسي',
                notes: ''
            };

            scholarships.unshift(newObj);
            saveToStorage();
            renderTable();
            updateStats();
            closeAddModal();
            e.target.reset();
            showToast('تمت إضافة المنحة الجديدة بنجاح إلى القائمة!', '🎉');
        }

        // Theme Toggle
        function toggleTheme() {
            const body = document.body;
            const current = body.getAttribute('data-theme');
            const next = current === 'dark' ? 'light' : 'dark';
            body.setAttribute('data-theme', next);
            document.getElementById('themeIcon').innerText = next === 'dark' ? '☀️' : '🌙';
            localStorage.setItem('radwa_scholarships_theme', next);
        }

        // Toast Helper
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
                "الرابط الرسمي": item.link,
                "الشروط الأساسية والمتطلبات": item.requirements,
                "ميعاد بداية التقديم": item.start_date,
                "آخر موعد للتقديم": item.deadline,
                "نوع التمويل والملاحظات": item.funding + (item.notes ? " | " + item.notes : "")
            }));

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "المنح والمؤسسات");
            XLSX.writeFile(wb, "Scholarships_Radwa_Exported.xlsx");
            showToast('تم تصدير ملف الإكسيل بنجاح!', '📊');
        }

        // Init on load - Default to Light Mode
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
