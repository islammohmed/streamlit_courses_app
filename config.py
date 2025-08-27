# Configuration for Streamlit Training Courses Management System

import os

## Application Settings
APP_TITLE = "نظام إدارة الدورات التدريبية"
APP_ICON = "📚"
PAGE_LAYOUT = "wide"

## File paths
SAMPLE_DATA_DIR = "sample_data"
EXCEL_FILE_PATH = "sample_data/النموذج-الموحد2025م.xlsx"
TEMPLATE_FILE_PATH = os.path.join(SAMPLE_DATA_DIR, "نموذج-اعتماد2.docx")

## Date Format Settings
DATE_FORMAT = "%Y-%m-%d"
DISPLAY_DATE_FORMAT = "%d/%m/%Y"

## Excel Column Mappings
# These are the expected column names in Arabic
EXCEL_COLUMNS = {
    "course_code": ["كود_الدورة", "رقم_الدورة", "كود الدورة"],
    "course_name": ["اسم_البرنامج", "اسم_الدورة", "اسم البرنامج", "اسم الدورة"],
    "trainer": ["المدرب", "اسم_المدرب", "اسم المدرب"],
    "target_audience": ["الجمهور_المستهدف", "الجمهور المستهدف", "الفئة_المستهدفة"],
    "start_date": ["تاريخ_البداية", "تاريخ البداية", "تاريخ_البدء", "تاريخ البدء"],
    "end_date": ["تاريخ_النهاية", "تاريخ النهاية", "تاريخ_الانتهاء", "تاريخ الانتهاء"],
    "status": ["حالة_الدورة", "حالة الدورة", "الحالة", "وضع_الدورة"],
    "participants": ["عدد_المتدربين", "عدد المتدربين", "عدد_المشاركين"],
    "hours": ["عدد_الساعات", "عدد الساعات", "ساعات_التدريب"],
    "location": ["المكان", "مكان_التدريب", "مكان التدريب", "القاعة"],
    "fees": ["الرسوم", "التكلفة", "السعر"],
    "notes": ["ملاحظات", "تعليقات", "ملاحظات_إضافية"]
}

## Status Values
STATUS_VALUES = {
    "executed": ["منفذة", "مكتملة", "منجزة", "executed", "completed"],
    "cancelled": ["ملغاة", "ملغية", "محذوفة", "cancelled", "canceled"],
    "postponed": ["مؤجلة", "مرجأة", "postponed", "delayed"],
    "planned": ["مخططة", "مجدولة", "planned", "scheduled"]
}

## Word Template Placeholders
# Common placeholders that might be used in Word templates
WORD_PLACEHOLDERS = [
    "اسم_المتدرب", "اسم_البرنامج", "كود_الدورة", "المدرب",
    "الجمهور_المستهدف", "تاريخ_البداية", "تاريخ_النهاية",
    "عدد_المتدربين", "عدد_الساعات", "المكان", "ملاحظات",
    "تاريخ_الاصدار", "الرسوم", "المؤسسة", "رقم_الشهادة"
]

## UI Text (Arabic Interface)
UI_TEXT = {
    "tabs": {
        "dashboard": "📊 لوحة الإحصائيات",
        "generator": "📄 توليد النماذج", 
        "comparison": "🔍 المقارنة"
    },
    "sidebar": {
        "settings": "⚙️ الإعدادات",
        "upload_excel": "رفع ملف Excel",
        "upload_template": "رفع قالب Word",
        "export_options": "📤 خيارات التصدير"
    },
    "dashboard": {
        "daily_stats": "📅 إحصائيات يومية",
        "monthly_stats": "📈 إحصائيات شهرية",
        "ongoing_courses": "دورات جارية",
        "starting_today": "دورات تبدأ اليوم",
        "ending_today": "دورات تنتهي اليوم",
        "cancelled_today": "دورات ملغاة"
    },
    "buttons": {
        "generate_form": "⬇️ توليد النموذج",
        "download_word": "📄 تحميل Word",
        "download_pdf": "📄 تحميل PDF",
        "bulk_generate": "توليد جميع النماذج وتحميل ملف ZIP",
        "export_comparison": "📄 تصدير تقرير المقارنة"
    }
}

## Chart Colors
CHART_COLORS = {
    "executed": "#28a745",    # Green
    "cancelled": "#dc3545",   # Red  
    "postponed": "#ffc107",   # Yellow
    "planned": "#007bff"      # Blue
}

## File Size Limits (in MB)
MAX_EXCEL_SIZE = 50
MAX_TEMPLATE_SIZE = 10

## Pagination Settings
DEFAULT_ITEMS_PER_PAGE = 10
ITEMS_PER_PAGE_OPTIONS = [5, 10, 20, 50]
