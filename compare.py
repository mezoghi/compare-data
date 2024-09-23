import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import time
import os

# وظيفة لاختيار ملف
def choose_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

# قراءة ملف Excel
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# مقارنة البيانات بين ملفين Excel
def compare_dataframes(df1, df2, key_column, compare_column):
    # الاحتفاظ فقط بالعمود المفتاح والعمود المطلوب
    df1_filtered = df1[[key_column, compare_column]].copy()
    df2_filtered = df2[[key_column, compare_column]].copy()

    # إعادة تسمية الأعمدة لتفادي التضارب
    df1_filtered = df1_filtered.rename(columns={compare_column: f"{compare_column}_file1"})
    df2_filtered = df2_filtered.rename(columns={compare_column: f"{compare_column}_file2"})

    # دمج الملفين على أساس العمود المفتاح
    merged = pd.merge(df1_filtered, df2_filtered, on=key_column, how='outer')

    # استخراج الاختلافات فقط
    differences = merged[merged[f"{compare_column}_file1"] != merged[f"{compare_column}_file2"]]

    # إرجاع الأعمدة المطلوبة فقط: العمود المفتاح + العمود المطلوب من الملفين
    return differences[[key_column, f"{compare_column}_file1", f"{compare_column}_file2"]]

# تحديث شريط التحميل
def update_progress_bar(current_step, total_steps):
    progress_value = (current_step / total_steps) * 100
    progress_bar['value'] = progress_value
    root.update_idletasks()  # تحديث الواجهة لعرض شريط التحميل

# تنفيذ عملية المقارنة
def compare_files():
    file1_path = entry_file1.get()
    file2_path = entry_file2.get()
    key_column = entry_key_column.get()
    compare_column = entry_compare_column.get()

    if not file1_path or not file2_path or not key_column or not compare_column:
        messagebox.showerror("خطأ", "يرجى اختيار كلا الملفين واسم العمود المفتاح والعمود المطلوب.")
        return

    try:
        if file1_path.endswith('.xlsx') and file2_path.endswith('.xlsx'):
            total_steps = 3  # إجمالي الخطوات لتحديث شريط التحميل (قراءة الملفين + المقارنة)
            progress_bar['value'] = 0

            # الخطوة 1: قراءة الملف الأول
            df1 = read_excel(file1_path)
            time.sleep(1)  # تأخير صغير لتحديث شريط التحميل
            update_progress_bar(1, total_steps)

            # الخطوة 2: قراءة الملف الثاني
            df2 = read_excel(file2_path)
            time.sleep(1)
            update_progress_bar(2, total_steps)

            # الخطوة 3: المقارنة بين الملفين
            differences = compare_dataframes(df1, df2, key_column, compare_column)
            update_progress_bar(3, total_steps)

            if differences.empty:
                messagebox.showinfo("نتيجة المقارنة", "لا توجد اختلافات في العمود المحدد.")
            else:
                # حفظ النتائج في ملف Excel
                output_file = os.path.join(os.path.expanduser("-"), "comparison_results.csv")
                differences.to_csv(output_file, index=False)
                messagebox.showinfo("نجاح", f"تم حفظ الاختلافات في الملف: {output_file}")
        else:
            messagebox.showerror("خطأ", "فقط ملفات Excel مدعومة حاليًا.")
    except Exception as e:
        progress_bar.stop()
        messagebox.showerror("خطأ", f"حدث خطأ أثناء المقارنة: {str(e)}")

# إعداد الواجهة
root = tk.Tk()
root.title("مقارنة الملفات")

# تخطيط الواجهة
frame = tk.Frame(root)
frame.pack(pady=20, padx=20)

# مدخل الملف الأول
label_file1 = tk.Label(frame, text="اختر الملف الأول:")
label_file1.grid(row=0, column=0, pady=5, padx=5, sticky="e")
entry_file1 = tk.Entry(frame, width=50)
entry_file1.grid(row=0, column=1, pady=5, padx=5)
button_file1 = tk.Button(frame, text="اختر الملف", command=lambda: choose_file(entry_file1))
button_file1.grid(row=0, column=2, pady=5, padx=5)

# مدخل الملف الثاني
label_file2 = tk.Label(frame, text="اختر الملف الثاني:")
label_file2.grid(row=1, column=0, pady=5, padx=5, sticky="e")
entry_file2 = tk.Entry(frame, width=50)
entry_file2.grid(row=1, column=1, pady=5, padx=5)
button_file2 = tk.Button(frame, text="اختر الملف", command=lambda: choose_file(entry_file2))
button_file2.grid(row=1, column=2, pady=5, padx=5)

# مدخل اسم العمود المفتاح
label_key_column = tk.Label(frame, text="ادخل اسم العمود المفتاح للمقارنة:")
label_key_column.grid(row=2, column=0, pady=5, padx=5, sticky="e")
entry_key_column = tk.Entry(frame, width=50)
entry_key_column.grid(row=2, column=1, pady=5, padx=5)

# مدخل اسم العمود المطلوب
label_compare_column = tk.Label(frame, text="ادخل اسم العمود المطلوب للمقارنة:")
label_compare_column.grid(row=3, column=0, pady=5, padx=5, sticky="e")
entry_compare_column = tk.Entry(frame, width=50)
entry_compare_column.grid(row=3, column=1, pady=5, padx=5)

# زر للمقارنة
button_compare = tk.Button(frame, text="قارن الملفات", command=compare_files)
button_compare.grid(row=4, column=1, pady=20)

# شريط التحميل
progress_bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate', length=400)
progress_bar.grid(row=5, column=1, pady=10)

# تشغيل الواجهة
root.mainloop()
