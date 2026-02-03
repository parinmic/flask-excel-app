from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = 'data.xlsx'
PRODUCT_FILE = 'products.xlsx'

def save_to_excel(data, excel_file=EXCEL_FILE):
    columns = ['ชื่อ', 'อีเมล', 'ข้อความ']
    new_df = pd.DataFrame([data], columns=columns)
    if not os.path.exists(excel_file):
        new_df.to_excel(excel_file, index=False)
    else:
        old_df = pd.read_excel(excel_file)
        df = pd.concat([old_df, new_df], ignore_index=True)
        df.to_excel(excel_file, index=False)

def save_product_to_excel(data):
    columns = ['รหัสสินค้า', 'รายการสินค้า']
    new_df = pd.DataFrame([data], columns=columns)
    if not os.path.exists(PRODUCT_FILE):
        new_df.to_excel(PRODUCT_FILE, index=False)
    else:
        old_df = pd.read_excel(PRODUCT_FILE)
        df = pd.concat([old_df, new_df], ignore_index=True)
        df.to_excel(PRODUCT_FILE, index=False)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        message = request.form['message']
        save_to_excel({'ชื่อ': name, 'อีเมล': email, 'ข้อความ': message})
        return redirect(url_for('index'))
    return render_template('form.html')

@app.route('/product', methods=['GET', 'POST'])
def product():
    success = False
    
    # อ่านข้อมูลจาก STOCK.xlsx
    products = []
    if os.path.exists('STOCK.xlsx'):
        df = pd.read_excel('STOCK.xlsx', header=None, skiprows=3)
        df.columns = ['col0', 'col1', 'รหัสสินค้า', 'รายการสินค้า', 
                      'ถุง_1วัน', 'นน_1วัน', 'ถุง_2วัน', 'นน_2วัน', 
                      'ถุง_3วัน', 'นน_3วัน', 'ถุง_3วัน+', 'นน_3วัน+']
        df = df[['รหัสสินค้า', 'รายการสินค้า', 'ถุง_1วัน', 'นน_1วัน', 
                 'ถุง_2วัน', 'นน_2วัน', 'ถุง_3วัน', 'นน_3วัน', 'ถุง_3วัน+', 'นน_3วัน+']]
        df = df.dropna(subset=['รหัสสินค้า'])
        df = df.fillna(0)
        products = df.to_dict('records')
    
    if request.method == 'POST':
        # บันทึกข้อมูลทั้งหมดลง Excel (รวมข้อมูลเดิมและตรวจสอบ)
        save_data = []
        for i, prod in enumerate(products, 1):
            check_1 = request.form.get(f'check_1_{i}', '')
            check_2 = request.form.get(f'check_2_{i}', '')
            check_3 = request.form.get(f'check_3_{i}', '')
            check_4 = request.form.get(f'check_4_{i}', '')
            
            save_data.append({
                'ลำดับ': i,
                'รหัสสินค้า': prod['รหัสสินค้า'],
                'รายการสินค้า': prod['รายการสินค้า'],
                'ถุง_1วัน': prod['ถุง_1วัน'],
                'นน_1วัน': prod['นน_1วัน'],
                'ตรวจสอบ_1วัน': check_1,
                'ถุง_2วัน': prod['ถุง_2วัน'],
                'นน_2วัน': prod['นน_2วัน'],
                'ตรวจสอบ_2วัน': check_2,
                'ถุง_3วัน': prod['ถุง_3วัน'],
                'นน_3วัน': prod['นน_3วัน'],
                'ตรวจสอบ_3วัน': check_3,
                'ถุง_>3วัน': prod['ถุง_3วัน+'],
                'นน_>3วัน': prod['นน_3วัน+'],
                'ตรวจสอบ_>3วัน': check_4
            })
        
        if save_data:
            save_df = pd.DataFrame(save_data)
            # สร้างชื่อไฟล์พร้อมวันที่และเวลา
            now = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'check_data_{now}.xlsx'
            save_df.to_excel(filename, index=False)
            success = True
    
    return render_template('product.html', products=products, success=success)

if __name__ == '__main__':
    # host='0.0.0.0' เพื่อให้เข้าถึงจากอุปกรณ์อื่นได้
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
