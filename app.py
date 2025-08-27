import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import openpyxl
from docx import Document
from docx.shared import Inches
import re
import io
import zipfile
import os
from pathlib import Path
import tempfile
import warnings
warnings.filterwarnings('ignore')

# Import configuration
import config

# Try to import docx2pdf for PDF generation
try:
    from docx2pdf import convert
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="نظام إدارة الدورات التدريبية",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# RTL CSS styling
st.markdown("""
<style>
    .main {
        direction: rtl;
        text-align: right;
    }
    .stSelectbox > div > div {
        direction: rtl;
    }
    .stDataFrame {
        direction: rtl;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin: 0.5rem 0;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #666;
    }
    .status-match {
        color: #28a745;
        font-weight: bold;
    }
    .status-missing {
        color: #dc3545;
        font-weight: bold;
    }
    .status-unused {
        color: #ffc107;
        font-weight: bold;
    }
    h1, h2, h3 {
        text-align: right;
        direction: rtl;
    }
    .stButton > button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ========================= DATA LOADING FUNCTIONS =========================

@st.cache_data
def load_excel_data(file_path, month_sheet=None):
    """
    Load Excel data from specified file and sheet
    Returns DataFrame with course data
    """
    try:
        if month_sheet:
            df = pd.read_excel(file_path, sheet_name=month_sheet)
        else:
            # Load first sheet if no specific sheet mentioned
            df = pd.read_excel(file_path)
        
        # Clean the data - remove empty rows and unnecessary columns
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Remove columns that are completely empty or just "Unnamed"
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # Filter rows that have actual course data (should have course name)
        course_name_cols = [col for col in df.columns if 'اسم الدورة' in str(col) and 'تاريخ' not in str(col)]
        if not course_name_cols:
            # Fallback: look for any column with 'اسم' and 'برنامج' or just 'اسم' and has substantial data
            course_name_cols = [col for col in df.columns if 'اسم' in str(col) and ('برنامج' in str(col) or 'دورة' in str(col)) and 'تاريخ' not in str(col)]
        
        if course_name_cols:
            course_col = course_name_cols[0]
            # Keep only rows that have course names
            df = df[df[course_col].notna() & (df[course_col] != '') & (df[course_col].astype(str).str.strip() != '')]
        else:
            # If no clear course column, just remove obviously empty rows
            # Keep rows where at least 3 columns have data
            df = df[df.count(axis=1) >= 3]
        
        # Ensure date columns are properly parsed
        date_columns = [col for col in df.columns if 'تاريخ' in str(col).lower() or 'date' in str(col).lower()]
        for col in date_columns:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            except:
                pass
        
        # Reset index after filtering
        df = df.reset_index(drop=True)
        
        return df
    except Exception as e:
        st.error(f"خطأ في تحميل البيانات: {str(e)}")
        return pd.DataFrame()

def get_available_sheets(file_path):
    """
    Get list of available sheets in Excel file
    """
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        return sheets
    except Exception as e:
        st.error(f"خطأ في قراءة أوراق العمل: {str(e)}")
        return []

# ========================= WORD TEMPLATE FUNCTIONS =========================

def extract_placeholders_from_word(docx_path):
    """
    Extract all placeholders from Word document
    Returns list of placeholders found
    """
    try:
        doc = Document(docx_path)
        placeholders = set()
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            matches = re.findall(r'\{\{([^}]+)\}\}', para.text)
            placeholders.update(matches)
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    matches = re.findall(r'\{\{([^}]+)\}\}', cell.text)
                    placeholders.update(matches)
        
        return list(placeholders)
    except Exception as e:
        st.error(f"خطأ في قراءة قالب Word: {str(e)}")
        return []

def build_mapping(row, df_columns):
    """
    Build mapping between Excel columns and Word Content Control tags
    Content Control Tags available:
    - اسم البرنامج, اسم الدورة, الفئة المستهدفة, طريقة الطرح
    - العدد المتوقع, وقت الدورة/البرنامج, تاريخ التنفيذ, مدتها
    - اســــم الــمــدرب, الـــمـتـطــلــبــات, تنفيذ البرنامج/الدورة, مقر التنفيذ
    """
    mapping = {}
    
    # Define Excel to Content Control tag mappings
    excel_to_tag_mapping = {
        # Excel Column Name -> Content Control Tag
        'اسم الدورة بالعربي': 'اسم الدورة',
        'الفئة المستهدفة': 'الفئة المستهدفة',
        'طريقة الطرح': 'طريقة الطرح',
        'اسم المدرب': 'اســــم الــمــدرب',
        'مكان الانعقاد ': 'مقر التنفيذ',  # Fixed: exact match with trailing space
        'مكان الانعقاد': 'مقر التنفيذ',   # Alternative without space
        'الوقت': 'وقت الدورة/البرنامج',
        'عدد الايام': 'مدتها',
        'تاريخ بداية الدورة بالميلادي': 'تاريخ التنفيذ',
        'تاريخ بداية الدورة بالهجري': 'تاريخ التنفيذ',
        'تحتاج لمعمل؟': 'استخدام معمل الحاسب',  # Fixed: lab field mapping
        
        # Additional mappings for other available fields
        'جهة التدريب': 'تنفيذ البرنامج/الدورة',
        'اسم جهة التدريب': 'تنفيذ البرنامج/الدورة',
    }
    
    # Build mapping from Excel data to Content Control tags
    st.write("🔍 **Debug: Mapping Process**")
    
    for excel_col, tag_name in excel_to_tag_mapping.items():
        value = ""
        
        # Try exact match first
        if excel_col in df_columns:
            value = str(row[excel_col]) if pd.notna(row[excel_col]) else ""
            st.write(f"✅ Found exact match: '{excel_col}' -> '{value}'")
        else:
            # Try similar column names (with space variations)
            excel_col_clean = excel_col.strip().lower()
            found_match = False
            
            # Special handling for location field with extensive matching
            if 'مكان الانعقاد' in excel_col:
                st.write(f"🔍 Looking for location field '{excel_col}'...")
                for col in df_columns:
                    # Check for exact match (with or without trailing space)
                    if col == 'مكان الانعقاد ' or col == 'مكان الانعقاد':
                        value = str(row[col]) if pd.notna(row[col]) else ""
                        st.write(f"🎯 Found exact location match: '{col}' -> '{value}'")
                        found_match = True
                        break
                    # Also check if column contains the location keywords
                    elif 'مكان' in str(col) and 'انعقاد' in str(col):
                        value = str(row[col]) if pd.notna(row[col]) else ""
                        st.write(f"🎯 Found location match: '{col}' -> '{value}'")
                        found_match = True
                        break
                
                if not found_match:
                    st.write(f"❌ No location field found. Available columns with 'مكان':")
                    location_cols = [col for col in df_columns if 'مكان' in str(col)]
                    for loc_col in location_cols:
                        st.write(f"   - '{loc_col}' (length: {len(str(loc_col))})")
            
            # Special handling for lab field
            elif 'تحتاج لمعمل' in excel_col or 'معمل' in excel_col:
                st.write(f"🔍 Looking for lab field '{excel_col}'...")
                for col in df_columns:
                    if col == 'تحتاج لمعمل؟' or 'تحتاج لمعمل' in str(col):
                        value = str(row[col]) if pd.notna(row[col]) else ""
                        st.write(f"🎯 Found exact lab match: '{col}' -> '{value}'")
                        found_match = True
                        break
                
                if not found_match:
                    st.write(f"❌ No lab field found. Available columns with 'معمل':")
                    lab_cols = [col for col in df_columns if 'معمل' in str(col)]
                    for lab_col in lab_cols:
                        st.write(f"   - '{lab_col}'")
            else:
                # Regular matching for other fields
                for col in df_columns:
                    col_clean = str(col).strip().lower()
                    if excel_col_clean == col_clean or excel_col_clean in col_clean or col_clean in excel_col_clean:
                        value = str(row[col]) if pd.notna(row[col]) else ""
                        st.write(f"🔍 Found similar match: '{col}' -> '{value}'")
                        found_match = True
                        break
        
        # Clean up the value
        if value and str(value).lower() not in ['nan', 'none', 'nat']:
            value = str(value).strip()
        else:
            value = ""
        
        # Map to Content Control tag name (only if we have data)
        if value:
            mapping[tag_name] = value
            st.write(f"✅ **MAPPED**: '{tag_name}' = '{value}'")
        else:
            st.write(f"❌ **NO DATA**: '{excel_col}' -> '{tag_name}'")
    
    return mapping

def generate_docx_from_template(template_path, mapping, output_name):
    """
    Generate Word document from template using Content Controls
    With complete error recovery and safe table handling
    """
    try:
        st.write("📄 Starting Word document generation...")
        
        # Load template with extreme caution
        try:
            doc = Document(template_path)
            st.write("✅ Template loaded successfully")
        except Exception as load_error:
            st.error(f"❌ خطأ في تحميل القالب: {str(load_error)}")
            return None
        
        # Debug: Show what mapping we received
        st.write("🔧 **Debug - Final Mapping**:")
        for key, value in mapping.items():
            st.write(f"  **{key}** -> '{value}'")
        
        replacements_made = 0
        
        # Method 1: Safe Content Controls processing
        st.write("🏷️ **Method 1: Processing Content Controls**")
        try:
            # Use safe iteration to avoid table errors
            content_controls_found = []
            
            # Find all content controls first without processing tables
            try:
                for element in doc.element.body.iter():
                    if element.tag.endswith('sdt'):  # Structured Document Tag
                        try:
                            tag_element = element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag')
                            if tag_element is not None:
                                tag_value = tag_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                if tag_value:
                                    content_controls_found.append((element, tag_value))
                        except Exception as cc_find_error:
                            continue
                
                st.write(f"🔍 Found {len(content_controls_found)} Content Controls")
                
                # Now process the found content controls
                for element, tag_value in content_controls_found:
                    try:
                        if tag_value.strip() in mapping:
                            data_value = mapping[tag_value.strip()]
                            
                            if data_value:
                                # Find text elements within the content control
                                text_elements = element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                                
                                if text_elements:
                                    # Clear existing text
                                    for text_elem in text_elements:
                                        text_elem.text = ""
                                    
                                    # Set new text (data only)
                                    if len(text_elements) > 0:
                                        text_elements[0].text = data_value
                                        replacements_made += 1
                                        st.write(f"✅ Updated '{tag_value}' with: '{data_value}'")
                    except Exception as cc_update_error:
                        st.warning(f"⚠️ Error updating content control '{tag_value}': {str(cc_update_error)}")
                        continue
                        
            except Exception as cc_method_error:
                st.warning(f"⚠️ Content Controls method failed: {str(cc_method_error)}")
                st.write("🔄 Switching to fallback method...")
                
                # Method 2: Safe text replacement fallback (avoid tables completely)
                st.write("📝 **Method 2: Safe Text Replacement**")
                try:
                    # Only process paragraphs, avoid tables to prevent grid_offset error
                    for paragraph in doc.paragraphs:
                        try:
                            original_text = paragraph.text
                            for tag_name, data_value in mapping.items():
                                if tag_name in paragraph.text and data_value:
                                    paragraph.text = paragraph.text.replace(tag_name, data_value)
                                    replacements_made += 1
                                    st.write(f"✅ Replaced '{tag_name}' in paragraph with: '{data_value}'")
                        except Exception as para_error:
                            continue
                            
                except Exception as fallback_error:
                    st.error(f"❌ All methods failed: {str(fallback_error)}")
        
        except Exception as processing_error:
            st.error(f"❌ Error in document processing: {str(processing_error)}")
        
        st.success(f"✅ **Total updates made: {replacements_made}**")
        
        # Save to BytesIO
        try:
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            st.write("💾 Document saved successfully")
            return doc_buffer
        except Exception as save_error:
            st.error(f"❌ Error saving document: {str(save_error)}")
            return None
        
    except Exception as e:
        st.error(f"❌ خطأ عام في توليد المستند: {str(e)}")
        st.write("**تفاصيل الخطأ:**")
        import traceback
        st.code(traceback.format_exc())
        return None

def convert_docx_to_pdf(docx_buffer, output_name):
    """
    Convert DOCX to PDF using docx2pdf
    """
    if not PDF_AVAILABLE:
        return None
    
    try:
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            temp_docx.write(docx_buffer.getvalue())
            temp_docx_path = temp_docx.name
        
        temp_pdf_path = temp_docx_path.replace('.docx', '.pdf')
        
        # Convert to PDF
        convert(temp_docx_path, temp_pdf_path)
        
        # Read PDF content
        with open(temp_pdf_path, 'rb') as pdf_file:
            pdf_content = pdf_file.read()
        
        # Cleanup
        os.unlink(temp_docx_path)
        os.unlink(temp_pdf_path)
        
        return pdf_content
    except Exception as e:
        st.warning(f"فشل في تحويل PDF: {str(e)}")
        return None

# ========================= DASHBOARD FUNCTIONS =========================

def calculate_daily_stats(df, target_date):
    """
    Calculate daily statistics for given date
    """
    stats = {
        'ongoing_courses': 0,
        'starting_today': 0,
        'ending_today': 0,
        'cancelled_today': 0
    }
    
    if df.empty:
        return stats
    
    # Find date columns
    start_date_col = None
    end_date_col = None
    status_col = None
    
    for col in df.columns:
        col_str = str(col).lower()
        if 'بداية' in col_str or 'start' in col_str:
            start_date_col = col
        elif 'نهاية' in col_str or 'end' in col_str:
            end_date_col = col
        elif 'حالة' in col_str or 'status' in col_str or 'وضع' in col_str:
            status_col = col
    
    if start_date_col and end_date_col:
        for _, row in df.iterrows():
            start_date = pd.to_datetime(row[start_date_col], errors='coerce')
            end_date = pd.to_datetime(row[end_date_col], errors='coerce')
            status = str(row[status_col]).lower() if status_col else ""
            
            if pd.notna(start_date) and pd.notna(end_date):
                # Ongoing courses
                if start_date.date() <= target_date <= end_date.date():
                    stats['ongoing_courses'] += 1
                
                # Starting today
                if start_date.date() == target_date:
                    stats['starting_today'] += 1
                
                # Ending today
                if end_date.date() == target_date:
                    stats['ending_today'] += 1
                
                # Cancelled today (if status indicates cancellation)
                if 'ملغ' in status or 'cancel' in status:
                    stats['cancelled_today'] += 1
    
    return stats

def calculate_monthly_stats(df):
    """
    Calculate monthly statistics
    """
    stats = {
        'total_planned': len(df),
        'executed': 0,
        'cancelled': 0,
        'postponed': 0,
        'total_training_days': 0
    }
    
    if df.empty:
        return stats
    
    # Find relevant columns
    status_col = None
    start_date_col = None
    end_date_col = None
    
    for col in df.columns:
        col_str = str(col).lower()
        if 'حالة' in col_str or 'status' in col_str or 'وضع' in col_str:
            status_col = col
        elif 'بداية' in col_str or 'start' in col_str:
            start_date_col = col
        elif 'نهاية' in col_str or 'end' in col_str:
            end_date_col = col
    
    # Calculate status distribution
    if status_col:
        for _, row in df.iterrows():
            status = str(row[status_col]).lower()
            if 'منفذ' in status or 'executed' in status or 'مكتمل' in status:
                stats['executed'] += 1
            elif 'ملغ' in status or 'cancel' in status:
                stats['cancelled'] += 1
            elif 'مؤجل' in status or 'postpone' in status:
                stats['postponed'] += 1
    
    # Calculate total training days
    if start_date_col and end_date_col:
        for _, row in df.iterrows():
            start_date = pd.to_datetime(row[start_date_col], errors='coerce')
            end_date = pd.to_datetime(row[end_date_col], errors='coerce')
            
            if pd.notna(start_date) and pd.notna(end_date):
                days = (end_date - start_date).days + 1
                stats['total_training_days'] += max(0, days)
    
    return stats

def build_dashboard(df):
    """
    Build the main dashboard with statistics and visualizations
    """
    st.header("📊 لوحة الإحصائيات")
    
    # Date selection
    col1, col2 = st.columns(2)
    with col1:
        selected_date = st.date_input("اختر التاريخ", datetime.now().date())
    with col2:
        if not df.empty:
            target_audience_col = None
            for col in df.columns:
                if 'جمهور' in str(col) or 'audience' in str(col).lower() or 'فئة' in str(col):
                    target_audience_col = col
                    break
            
            if target_audience_col:
                audiences = ['الكل'] + list(df[target_audience_col].dropna().unique())
                selected_audience = st.selectbox("فلترة حسب الجمهور المستهدف", audiences)
                
                if selected_audience != 'الكل':
                    df = df[df[target_audience_col] == selected_audience]
    
    # Daily Statistics
    st.subheader("📅 إحصائيات يومية")
    daily_stats = calculate_daily_stats(df, selected_date)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{daily_stats['ongoing_courses']}</div>
            <div class="metric-label">دورات جارية</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{daily_stats['starting_today']}</div>
            <div class="metric-label">دورات تبدأ اليوم</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{daily_stats['ending_today']}</div>
            <div class="metric-label">دورات تنتهي اليوم</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{daily_stats['cancelled_today']}</div>
            <div class="metric-label">دورات ملغاة</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Monthly Statistics
    st.subheader("📈 إحصائيات شهرية")
    monthly_stats = calculate_monthly_stats(df)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("إجمالي الدورات المخططة", monthly_stats['total_planned'])
    with col2:
        st.metric("دورات منفذة", monthly_stats['executed'])
    with col3:
        st.metric("دورات ملغاة", monthly_stats['cancelled'])
    with col4:
        st.metric("إجمالي أيام التدريب", monthly_stats['total_training_days'])
    
    # Visualizations
    if not df.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("توزيع حالة الدورات")
            status_data = {
                'الحالة': ['منفذة', 'ملغاة', 'مؤجلة'],
                'العدد': [monthly_stats['executed'], monthly_stats['cancelled'], monthly_stats['postponed']]
            }
            fig_bar = px.bar(status_data, x='الحالة', y='العدد', 
                           title="توزيع حالة الدورات",
                           color='الحالة',
                           color_discrete_map={'منفذة': '#28a745', 'ملغاة': '#dc3545', 'مؤجلة': '#ffc107'})
            fig_bar.update_layout(showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)
        
        with col2:
            st.subheader("عدد الدورات عبر الأيام")
            # Create sample data for courses per day
            if len(df) > 0:
                # Find start date column
                start_date_col = None
                for col in df.columns:
                    if 'بداية' in str(col).lower() or 'start' in str(col).lower():
                        start_date_col = col
                        break
                
                if start_date_col:
                    df_dates = df.copy()
                    df_dates[start_date_col] = pd.to_datetime(df_dates[start_date_col], errors='coerce')
                    df_dates = df_dates.dropna(subset=[start_date_col])
                    
                    if not df_dates.empty:
                        daily_courses = df_dates.groupby(df_dates[start_date_col].dt.date).size().reset_index()
                        daily_courses.columns = ['التاريخ', 'عدد الدورات']
                        
                        fig_line = px.line(daily_courses, x='التاريخ', y='عدد الدورات',
                                         title="عدد الدورات يومياً")
                        st.plotly_chart(fig_line, use_container_width=True)
    
    # Summary Table
    st.subheader("📋 ملخص الدورات الرئيسية")
    if not df.empty:
        # Show ongoing and cancelled courses
        summary_df = df.head(10)  # Show first 10 courses
        st.write(summary_df.to_html(escape=False), unsafe_allow_html=True)

# ========================= FORM GENERATOR FUNCTIONS =========================

def build_form_generator(df, template_path):
    """
    Build the accreditation form generator interface
    """
    st.header("📄 مولد نماذج الاعتماد")
    
    if df.empty:
        st.warning("لا توجد بيانات لعرضها")
        return
    
    if not template_path:
        st.warning("يرجى رفع قالب Word أولاً")
        return
    
    # Pagination settings
    items_per_page = st.selectbox("عدد العناصر في الصفحة", [5, 10, 20, 50], index=1)
    
    total_items = len(df)
    total_pages = (total_items - 1) // items_per_page + 1
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        current_page = st.number_input("الصفحة", min_value=1, max_value=total_pages, value=1)
    
    # Calculate start and end indices
    start_idx = (current_page - 1) * items_per_page
    end_idx = min(start_idx + items_per_page, total_items)
    
    # Display current page data
    page_df = df.iloc[start_idx:end_idx].copy()
    page_df.reset_index(drop=True, inplace=True)
    
    # Display data with generate buttons
    st.subheader(f"البيانات (الصفحة {current_page} من {total_pages})")
    
    # Extract placeholders from template
    placeholders = extract_placeholders_from_word(template_path)
    
    # Show data table with generate buttons
    for idx, row in page_df.iterrows():
        with st.expander(f"السطر {start_idx + idx + 1}"):
            col1, col2 = st.columns([3, 1])
            
            with col1:
                # Display row data
                row_dict = row.to_dict()
                for key, value in row_dict.items():
                    st.text(f"{key}: {value}")
            
            with col2:
                if st.button(f"⬇️ توليد النموذج", key=f"generate_{start_idx + idx}"):
                    # Generate mapping
                    mapping = build_mapping(row, df.columns.tolist())
                    
                    # Generate DOCX
                    output_name = f"نموذج_السطر_{start_idx + idx + 1}"
                    docx_buffer = generate_docx_from_template(template_path, mapping, output_name)
                    
                    if docx_buffer:
                        # Download DOCX
                        st.download_button(
                            label="📄 تحميل Word",
                            data=docx_buffer.getvalue(),
                            file_name=f"{output_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"download_docx_{start_idx + idx}"
                        )
                        
                        # Generate PDF if available
                        if PDF_AVAILABLE:
                            pdf_content = convert_docx_to_pdf(docx_buffer, output_name)
                            if pdf_content:
                                st.download_button(
                                    label="📄 تحميل PDF",
                                    data=pdf_content,
                                    file_name=f"{output_name}.pdf",
                                    mime="application/pdf",
                                    key=f"download_pdf_{start_idx + idx}"
                                )
    
    # Bulk generation option
    st.subheader("📦 التوليد المجمع")
    if st.button("توليد جميع النماذج وتحميل ملف ZIP"):
        with st.spinner("جاري توليد النماذج..."):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for idx, row in df.iterrows():
                    try:
                        # Generate mapping
                        mapping = build_mapping(row, df.columns.tolist())
                        
                        # Generate DOCX
                        output_name = f"نموذج_السطر_{idx + 1}"
                        docx_buffer = generate_docx_from_template(template_path, mapping, output_name)
                        
                        if docx_buffer:
                            zip_file.writestr(f"{output_name}.docx", docx_buffer.getvalue())
                            
                            # Add PDF if available
                            if PDF_AVAILABLE:
                                pdf_content = convert_docx_to_pdf(docx_buffer, output_name)
                                if pdf_content:
                                    zip_file.writestr(f"{output_name}.pdf", pdf_content)
                    except Exception as e:
                        st.warning(f"خطأ في توليد النموذج للسطر {idx + 1}: {str(e)}")
            
            zip_buffer.seek(0)
            
            st.download_button(
                label="📦 تحميل جميع النماذج (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="جميع_النماذج.zip",
                mime="application/zip"
            )

# ========================= COMPARISON FUNCTIONS =========================

def build_comparison_view(df, template_path):
    """
    Build Excel vs Word comparison interface
    """
    st.header("🔍 مقارنة Excel مع Word")
    
    if not template_path:
        st.warning("يرجى رفع قالب Word أولاً")
        return
    
    # Extract placeholders from Word template
    placeholders = extract_placeholders_from_word(template_path)
    
    if not placeholders:
        st.warning("لم يتم العثور على أي عناصر نائبة في قالب Word")
        return
    
    # Get Excel columns
    excel_columns = df.columns.tolist() if not df.empty else []
    
    # Build comparison data
    comparison_data = []
    matched_columns = set()
    
    for placeholder in placeholders:
        # Try to find exact match
        if placeholder in excel_columns:
            comparison_data.append({
                "العنصر النائب في Word": f"{{{{{placeholder}}}}}",
                "عمود Excel المطابق": placeholder,
                "الحالة": "✅ مطابق"
            })
            matched_columns.add(placeholder)
        else:
            # Try to find similar column
            similar_cols = [col for col in excel_columns if placeholder.lower() in str(col).lower() or str(col).lower() in placeholder.lower()]
            if similar_cols:
                comparison_data.append({
                    "العنصر النائب في Word": f"{{{{{placeholder}}}}}",
                    "عمود Excel المطابق": similar_cols[0],
                    "الحالة": "⚠️ مشابه"
                })
                matched_columns.add(similar_cols[0])
            else:
                comparison_data.append({
                    "العنصر النائب في Word": f"{{{{{placeholder}}}}}",
                    "عمود Excel المطابق": "غير موجود",
                    "الحالة": "❌ مفقود"
                })
    
    # Add unused Excel columns
    unused_columns = set(excel_columns) - matched_columns
    for col in unused_columns:
        comparison_data.append({
            "العنصر النائب في Word": "غير موجود",
            "عمود Excel المطابق": col,
            "الحالة": "⚠️ غير مستخدم"
        })
    
    # Display comparison table
    st.subheader("📊 جدول المقارنة")
    comparison_df = pd.DataFrame(comparison_data)
    
    # Apply styling based on status
    def style_status(val):
        if "✅" in val:
            return 'color: #28a745; font-weight: bold'
        elif "❌" in val:
            return 'color: #dc3545; font-weight: bold'
        elif "⚠️" in val:
            return 'color: #ffc107; font-weight: bold'
        return ''
    
    # Display the comparison table using HTML to avoid pyarrow dependency
    st.write(comparison_df.to_html(escape=False), unsafe_allow_html=True)
    
    # Summary statistics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        matched_count = len([item for item in comparison_data if "✅" in item["الحالة"]])
        st.metric("عناصر مطابقة", matched_count)
    
    with col2:
        similar_count = len([item for item in comparison_data if "⚠️" in item["الحالة"] and "مشابه" in item["الحالة"]])
        st.metric("عناصر مشابهة", similar_count)
    
    with col3:
        missing_count = len([item for item in comparison_data if "❌" in item["الحالة"]])
        st.metric("عناصر مفقودة", missing_count)
    
    with col4:
        unused_count = len([item for item in comparison_data if "غير مستخدم" in item["الحالة"]])
        st.metric("أعمدة غير مستخدمة", unused_count)
    
    # Warnings and recommendations
    if missing_count > 0:
        st.warning(f"⚠️ يوجد {missing_count} عنصر نائب في قالب Word لا يوجد له عمود مطابق في Excel")
    
    if unused_count > 0:
        st.info(f"💡 يوجد {unused_count} عمود في Excel لا يستخدم في قالب Word")
    
    # Export comparison report
    if st.button("📄 تصدير تقرير المقارنة"):
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            comparison_df.to_excel(writer, sheet_name='تقرير المقارنة', index=False)
        
        excel_buffer.seek(0)
        st.download_button(
            label="📊 تحميل تقرير Excel",
            data=excel_buffer.getvalue(),
            file_name="تقرير_المقارنة.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ========================= MAIN APPLICATION =========================

def main():
    """
    Main application function
    """
    st.title("📚 نظام إدارة الدورات التدريبية")
    st.markdown("---")
    
    # Sidebar for file uploads and settings
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        
        # Check for default files first with complete error handling
        try:
            if hasattr(config, 'EXCEL_FILE_PATH') and hasattr(config, 'TEMPLATE_FILE_PATH'):
                default_excel_path = config.EXCEL_FILE_PATH
                default_template_path = config.TEMPLATE_FILE_PATH
                
                # Test template loading early to catch errors
                try:
                    test_doc = Document(default_template_path)
                    st.success("✅ قالب Word تم تحميله بنجاح")
                except Exception as template_test_error:
                    st.error(f"❌ خطأ في قراءة قالب Word: {str(template_test_error)}")
                    default_template_path = None
            else:
                st.error("❌ خطأ في إعدادات الملفات")
                default_excel_path = None
                default_template_path = None
        except Exception as config_error:
            st.error(f"❌ خطأ في قراءة الإعدادات: {str(config_error)}")
            default_excel_path = None
            default_template_path = None
        
        excel_df = pd.DataFrame()
        available_sheets = []
        excel_path = None
        template_path = None
        
        # Try to load default Excel file if it exists
        if os.path.exists(default_excel_path):
            st.info(f"📊 تم العثور على ملف البيانات: {os.path.basename(default_excel_path)}")
            excel_path = default_excel_path
            available_sheets = get_available_sheets(excel_path)
            
            if available_sheets:
                # Auto-select September sheet if available, otherwise show selector
                september_sheet = None
                september_variations = ["سبتمبر", "September", "9", "09"]
                
                for variation in september_variations:
                    if variation in available_sheets:
                        september_sheet = variation
                        break
                
                if september_sheet:
                    st.success(f"🗓️ تم تحديد شهر سبتمبر تلقائياً: '{september_sheet}'")
                    selected_sheet = september_sheet
                    # Also show option to change if needed
                    if st.checkbox("تغيير الشهر", key="change_month"):
                        selected_sheet = st.selectbox("اختر الشهر (ورقة العمل)", available_sheets, 
                                                    index=available_sheets.index(september_sheet))
                else:
                    st.warning("لم يتم العثور على شهر سبتمبر، يرجى اختيار الشهر:")
                    selected_sheet = st.selectbox("اختر الشهر (ورقة العمل)", available_sheets)
                
                excel_df = load_excel_data(excel_path, selected_sheet)
                
                if not excel_df.empty:
                    st.success(f"تم تحميل {len(excel_df)} صف من بيانات شهر: {selected_sheet}")
                    
                    # Show data summary
                    st.info(f"📋 ملخص البيانات: {len(excel_df)} دورة في {selected_sheet}")
        else:
            st.warning("لم يتم العثور على ملف البيانات الافتراضي")
        
        # Excel file upload (alternative)
        st.markdown("**أو قم برفع ملف Excel آخر:**")
        excel_file = st.file_uploader("رفع ملف Excel", type=['xlsx', 'xls'])
        
        if excel_file:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(excel_file.getvalue())
                excel_path = tmp.name
            
            # Get available sheets
            available_sheets = get_available_sheets(excel_path)
            
            if available_sheets:
                # Auto-select September sheet if available
                september_sheet = None
                september_variations = ["سبتمبر", "September", "9", "09"]
                
                for variation in september_variations:
                    if variation in available_sheets:
                        september_sheet = variation
                        break
                
                if september_sheet:
                    st.success(f"🗓️ تم تحديد شهر سبتمبر تلقائياً: '{september_sheet}'")
                    default_index = available_sheets.index(september_sheet)
                else:
                    default_index = 0
                
                selected_sheet = st.selectbox("اختر الشهر (ورقة العمل)", available_sheets, 
                                            index=default_index, key="uploaded_sheet")
                excel_df = load_excel_data(excel_path, selected_sheet)
                
                if not excel_df.empty:
                    st.success(f"تم تحميل {len(excel_df)} صف من البيانات")
                    
                    # Show preview
                    st.subheader("📋 معاينة البيانات")
                    st.write(f"عدد الصفوف: {len(excel_df)}")
                    st.write(f"عدد الأعمدة: {len(excel_df.columns)}")
                    
                    # Display first few rows using HTML table to avoid pyarrow
                    html_table = excel_df.head().to_html(escape=False, index=False)
                    st.markdown("**أول 5 صفوف:**")
                    st.markdown(html_table, unsafe_allow_html=True)
            
            # Clean up temp file
            try:
                os.unlink(excel_path)
            except:
                pass
        
        # Check for default Word template
        if os.path.exists(default_template_path):
            st.info(f"📄 تم العثور على قالب Word: {os.path.basename(default_template_path)}")
            template_path = default_template_path
        
        # Word template upload (alternative)
        st.markdown("**أو قم برفع قالب Word آخر:**")
        template_file = st.file_uploader("رفع قالب Word", type=['docx'])
        
        if template_file:
            # Save uploaded template temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                tmp.write(template_file.getvalue())
                template_path = tmp.name
            
            st.success("تم رفع قالب Word بنجاح")
        
        # Export options
        st.markdown("---")
        st.subheader("📤 خيارات التصدير")
        
        if not excel_df.empty:
            # Monthly summary export
            if st.button("تصدير الملخص الشهري"):
                monthly_stats = calculate_monthly_stats(excel_df)
                
                summary_data = {
                    'المؤشر': ['إجمالي الدورات المخططة', 'دورات منفذة', 'دورات ملغاة', 'دورات مؤجلة', 'إجمالي أيام التدريب'],
                    'القيمة': [monthly_stats['total_planned'], monthly_stats['executed'], 
                              monthly_stats['cancelled'], monthly_stats['postponed'], monthly_stats['total_training_days']]
                }
                
                summary_df = pd.DataFrame(summary_data)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    summary_df.to_excel(writer, sheet_name='الملخص الشهري', index=False)
                    excel_df.to_excel(writer, sheet_name='البيانات الكاملة', index=False)
                
                excel_buffer.seek(0)
                st.download_button(
                    label="📊 تحميل الملخص الشهري",
                    data=excel_buffer.getvalue(),
                    file_name=f"الملخص_الشهري_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    # Main content tabs
    tab1, tab2, tab3 = st.tabs(["📊 لوحة الإحصائيات", "📄 توليد النماذج", "🔍 المقارنة"])
    
    with tab1:
        build_dashboard(excel_df)
    
    with tab2:
        build_form_generator(excel_df, template_path)
    
    with tab3:
        build_comparison_view(excel_df, template_path)
    
    # Cleanup temporary files (only if they were uploaded, not default files)
    if excel_file and excel_path and excel_path != config.EXCEL_FILE_PATH:
        try:
            os.unlink(excel_path)
        except:
            pass
    
    if template_file and template_path and template_path != config.TEMPLATE_FILE_PATH:
        try:
            os.unlink(template_path)
        except:
            pass

if __name__ == "__main__":
    main()
