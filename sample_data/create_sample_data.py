import pandas as pd
from datetime import datetime, timedelta
import random

# Create sample training courses data
def create_sample_data():
    # Sample data for training courses
    courses = []
    
    # Course names in Arabic
    course_names = [
        "دورة إدارة المشاريع",
        "دورة تطوير المهارات القيادية",
        "دورة الأمن السيبراني",
        "دورة تحليل البيانات",
        "دورة التسويق الرقمي",
        "دورة إدارة الموارد البشرية",
        "دورة خدمة العملاء",
        "دورة المحاسبة المالية",
        "دورة البرمجة والتطوير",
        "دورة إدارة الجودة"
    ]
    
    # Target audiences
    audiences = ["موظفو الحكومة", "القطاع الخاص", "الطلاب", "رجال الأعمال", "المهنيون"]
    
    # Statuses
    statuses = ["منفذة", "ملغاة", "مؤجلة", "مخططة"]
    
    # Generate data for current month
    current_date = datetime.now()
    start_of_month = current_date.replace(day=1)
    
    for i in range(50):  # Generate 50 courses
        start_date = start_of_month + timedelta(days=random.randint(0, 30))
        end_date = start_date + timedelta(days=random.randint(1, 5))
        
        course = {
            "كود_الدورة": f"TR{2024}{i+1:03d}",
            "اسم_البرنامج": random.choice(course_names),
            "اسم_الدورة": random.choice(course_names),
            "المدرب": f"د. أحمد محمد {i+1}",
            "الجمهور_المستهدف": random.choice(audiences),
            "تاريخ_البداية": start_date,
            "تاريخ_النهاية": end_date,
            "عدد_المتدربين": random.randint(10, 30),
            "عدد_الساعات": random.randint(8, 40),
            "المكان": f"قاعة التدريب {random.randint(1, 5)}",
            "حالة_الدورة": random.choice(statuses),
            "الرسوم": random.randint(500, 3000),
            "ملاحظات": "دورة تدريبية متخصصة"
        }
        courses.append(course)
    
    return pd.DataFrame(courses)

# Create and save sample data
if __name__ == "__main__":
    df = create_sample_data()
    
    # Save to Excel with multiple sheets (months)
    with pd.ExcelWriter('sample_courses.xlsx') as writer:
        df.to_excel(writer, sheet_name='يناير 2024', index=False)
        df.to_excel(writer, sheet_name='فبراير 2024', index=False)
        df.to_excel(writer, sheet_name='مارس 2024', index=False)
    
    print("Sample Excel file created: sample_courses.xlsx")
    print(f"Generated {len(df)} courses with columns: {list(df.columns)}")
