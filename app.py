# -*- coding: utf-8 -*-
import os
import smtplib
import socket
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# =========================================================================
# PythonAnywhere 免費版 IPv6 限制修復 (強制使用 IPv4 解析)
# =========================================================================
orig_getaddrinfo = socket.getaddrinfo
def getaddrinfo_ipv4(host, port, family=0, type=0, proto=0, flags=0):
    return orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
socket.getaddrinfo = getaddrinfo_ipv4
# =========================================================================

from flask import Flask, render_template, request, redirect, url_for, session
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = 'sweet_cake_shop_secret_key_for_session'

# =========================================================================
# 這裡要改成自己的 Email 設定 (SMTP 設定)
# =========================================================================
# 如果您使用的是 Gmail，請至 Google 帳戶啟用「雙重驗證」，並建立一個「應用程式密碼」填入下方。
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SENDER_EMAIL = 'chaozhiyi5464@gmail.com'     # 【這裡要改成自己的 Email】
SENDER_PASSWORD = 'grjrerzkobtbxzgj'     # 【這裡要改成自己的應用程式密碼】
# =========================================================================

# =========================================================================
# 管理員後台登入密碼
# =========================================================================
ADMIN_PASSWORD = '0146c016'                  # 【這裡可以修改後台登入密碼】
# =========================================================================


# Excel 存檔路徑，預設在專案目錄下
EXCEL_FILE = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'orders.xlsx')

def save_to_excel(order_data):
    """
    將訂單資料寫入 Excel (orders.xlsx)。
    如果檔案不存在則建立新檔並加上表頭；若已存在則追加至下一列。
    """
    try:
        if not os.path.exists(EXCEL_FILE):
            # 建立新的 Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "訂單資料"
            # 定義表頭欄位
            headers = [
                "姓名", "電話", "Email", "蛋糕品項", "數量", 
                "取貨或外送方式", "外送地址", "希望取貨或外送日期", 
                "備註", "訂購時間"
            ]
            ws.append(headers)
        else:
            # 載入現有 Workbook
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
        
        # 整理訂單列資料
        row = [
            order_data.get('name'),
            order_data.get('phone'),
            order_data.get('email'),
            order_data.get('cake_item'),
            int(order_data.get('quantity', 1)),
            order_data.get('delivery_type'),
            order_data.get('address', ''),
            order_data.get('delivery_date'),
            order_data.get('notes', ''),
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 訂購時間
        ]
        
        ws.append(row)
        wb.save(EXCEL_FILE)
        print(f"[系統] 訂單已成功寫入 Excel: {EXCEL_FILE}")
        return True
    except Exception as e:
        print(f"[錯誤] 寫入 Excel 失敗: {str(e)}")
        return False

def read_orders_from_excel():
    """
    從 Excel (orders.xlsx) 讀取所有訂單。
    用於後台網頁 (/admin) 顯示。
    """
    orders = []
    if not os.path.exists(EXCEL_FILE):
        return orders
    
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb.active
        # 從第二列開始讀取 (跳過表頭)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):  # 略過空列
                continue
            orders.append({
                'name': row[0],
                'phone': row[1],
                'email': row[2],
                'cake_item': row[3],
                'quantity': row[4],
                'delivery_type': row[5],
                'address': row[6] if row[6] else '無',
                'delivery_date': row[7],
                'notes': row[8] if row[8] else '無',
                'order_time': row[9]
            })
        # 將訂單反轉，讓最新的訂單排在最前面
        orders.reverse()
    except Exception as e:
        print(f"[錯誤] 讀取 Excel 失敗: {str(e)}")
    
    return orders

def send_notification_email(order_data):
    """
    自動發送「訂購成功通知信」給顧客。
    若寄信失敗或設定未修改，將會記錄在終端機而不影響訂單成立。
    """
    # 判斷是否為預設的 Email 占位符
    if SENDER_EMAIL == 'your_email@gmail.com' or SENDER_PASSWORD == 'your_app_password':
        print("\n=== [Email 測試模式] 偵測到尚未設定真實 Email，以下為模擬寄送之內容 ===")
        print(f"收件人 (顧客): {order_data['email']}")
        print(f"副本收件人 (店家/我): {SENDER_EMAIL}")
        print(f"主旨: 🧁 甜點時光 - 您的蛋糕訂單已送出成功！")
        print(f"顧客姓名: {order_data['name']}")
        print(f"訂購品項: {order_data['cake_item']} x {order_data['quantity']}")
        print(f"電話: {order_data['phone']}")
        print(f"取貨方式: {order_data['delivery_type']}")
        if order_data['delivery_type'] == '外送':
            print(f"外送地址: {order_data['address']}")
        print(f"希望日期: {order_data['delivery_date']}")
        print(f"備註: {order_data['notes'] or '無'}")
        print("===================================================================\n")
        return True, "Email 未設定，已模擬發送至顧客與您自己（請查看終端機）"

    try:
        # 建立郵件訊息
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"🧁 甜點時光 - 您的訂單已訂購成功！"
        msg['From'] = SENDER_EMAIL
        msg['To'] = order_data['email']
        msg['Cc'] = SENDER_EMAIL  # 在信件標頭加入副本

        # 設計精美溫馨的 HTML 電子郵件版面
        html_content = f"""
        <html>
        <body style="font-family: 'Microsoft JhengHei', sans-serif; color: #5D4037; background-color: #FFF5F5; padding: 20px; margin: 0;">
            <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 30px; border-radius: 15px; border: 2px solid #FFDAC1; box-shadow: 0 4px 15px rgba(255, 183, 178, 0.2);">
                <div style="text-align: center; border-bottom: 2px dashed #FFB7B2; padding-bottom: 15px; margin-bottom: 20px;">
                    <h1 style="color: #E05780; margin: 0; font-size: 24px;">🍰 甜點時光 Cake Shop</h1>
                    <p style="color: #8D6E63; margin: 5px 0 0 0; font-size: 14px;">用最溫暖的手作甜點，陪伴您的美好日常</p>
                </div>
                
                <p style="font-size: 16px; line-height: 1.6;">親愛的 <strong>{order_data['name']}</strong> 您好：</p>
                <p style="font-size: 16px; line-height: 1.6;">非常感謝您向我們訂購蛋糕！我們已經收到您的線上訂單，正在準備為您製作。以下是您的訂單詳細資訊：</p>
                
                <div style="background-color: #FFF5F5; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #FFB7B2;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 15px;">
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; width: 130px; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">訂購品項：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9; font-weight: bold; color: #E05780;">{order_data['cake_item']}</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">訂購數量：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9;">{order_data['quantity']} 個</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">聯絡電話：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9;">{order_data['phone']}</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">取貨或外送：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9;">{order_data['delivery_type']}</td>
                        </tr>
                        """
        
        # 如果是外送，則在信件中顯示外送地址
        if order_data['delivery_type'] == '外送':
            html_content += f"""
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">外送地址：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9;">{order_data['address']}</td>
                        </tr>
            """
            
        html_content += f"""
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #FFEBE9; color: #8D6E63;">希望日期：</td>
                            <td style="padding: 10px 0; border-bottom: 1px solid #FFEBE9;">{order_data['delivery_date']}</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px 0; font-weight: bold; color: #8D6E63;">備註說明：</td>
                            <td style="padding: 10px 0;">{order_data['notes'] if order_data['notes'] else '無'}</td>
                        </tr>
                    </table>
                </div>
                
                <div style="background-color: #FFF9C4; padding: 15px; border-radius: 8px; text-align: center; border: 1px dashed #FBC02D; margin: 25px 0;">
                    <p style="margin: 0; font-size: 15px; color: #F57F17; font-weight: bold;">
                        ✨ 訂購成功提醒 ✨
                    </p>
                    <p style="margin: 5px 0 0 0; font-size: 13px; color: #5D4037; line-height: 1.5;">
                        自取顧客請於指定日期前往門市領取；外送顧客請於當天保持電話暢通，外送夥伴會與您聯繫喔！
                    </p>
                </div>
                
                <p style="font-size: 15px; line-height: 1.6; text-align: center; color: #E05780; font-weight: bold;">
                    🧁 甜點新鮮製作，期待與您分享最甜蜜的手作幸福！ 🧁
                </p>
                
                <div style="text-align: center; margin-top: 35px; border-top: 1px solid #FFDAC1; padding-top: 15px; font-size: 12px; color: #8D6E63;">
                    <p style="margin: 0;">此信件為系統自動發送，請勿直接回覆。</p>
                    <p style="margin: 5px 0 0 0;">甜點時光門市電話：(02) 1234-5678 | 營業時間：11:00 - 20:00 (週一公休)</p>
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_content, 'html'))

        # 同時傳送給顧客與自己，用 set 去重防止填入相同信箱時重複傳送
        recipients = list(set([order_data['email'], SENDER_EMAIL]))

        # 發送 Email (先嘗試 Port 587，若失敗則嘗試 Port 465)
        try:
            print("[Email] 嘗試使用 Port 587 (STARTTLS) 連線...")
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10)
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            server.quit()
            print(f"[系統] 成功寄送訂單確認信至顧客: {order_data['email']} (Port 587)")
            return True, "Email 已成功寄出給您與顧客"
        except Exception as e587:
            print(f"[警告] Port 587 發送失敗: {str(e587)}。嘗試備用方案 Port 465 (SSL)...")
            try:
                server = smtplib.SMTP_SSL(SMTP_SERVER, 465, timeout=10)
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
                server.quit()
                print(f"[系統] 成功使用備用方案 (Port 465 SSL) 寄送電子郵件")
                return True, "Email 已成功寄出 (備用通道 465)"
            except Exception as e465:
                error_msg = f"Email 寄送失敗。Port 587 錯誤: {str(e587)} | Port 465 錯誤: {str(e465)}"
                print(f"[錯誤] {error_msg}")
                return False, error_msg
    except Exception as e_outer:
        error_msg = f"Email 建立或準備階段失敗: {str(e_outer)}"
        print(f"[錯誤] {error_msg}")
        return False, error_msg

# =========================================================================
# 路由設計 (Routing)
# =========================================================================

@app.route('/')
def index():
    """
    網站首頁：展示品牌介紹與人氣商品
    """
    return render_template('index.html')

@app.route('/order', methods=['GET', 'POST'])
def order():
    """
    訂購表單頁面
    GET: 渲染空白訂購表單
    POST: 接收顧客提交的表單資料，寫入 Excel 並發送 Email，最後導向成功頁面
    """
    if request.method == 'POST':
        # 接收表單資料
        order_data = {
            'name': request.form.get('name', '').strip(),
            'phone': request.form.get('phone', '').strip(),
            'email': request.form.get('email', '').strip(),
            'cake_item': request.form.get('cake_item', ''),
            'quantity': request.form.get('quantity', '1'),
            'delivery_type': request.form.get('delivery_type', '自取'),
            'address': request.form.get('address', '').strip(),
            'delivery_date': request.form.get('delivery_date', ''),
            'notes': request.form.get('notes', '').strip()
        }

        # 寫入 Excel
        excel_success = save_to_excel(order_data)
        
        # 寄送 Email (如果發送失敗，也會在後台印出，不影響網頁顯示)
        email_success, email_info = send_notification_email(order_data)
        
        # 將訂單資訊與寄信狀態存入 Session，以便在成功頁面顯示
        session['last_order'] = order_data
        session['email_info'] = email_info
        
        return redirect(url_for('success'))

    # 人氣蛋糕選項，方便在表單中生成下拉選單
    cakes = [
        "招牌鮮奶油草莓蛋糕",
        "布朗尼",
        "經典芋泥奶油蛋糕",
        "覆盆子蛋糕",
        "碧玉葡萄塔",
        "巧克力鮮奶油蛋糕",
        "藍莓塔",
        "冰心芒果捲",
        "法式檸檬塔",
        "黑絲絨巧克力蛋糕",
        "波士頓派",
        "海鹽雙層乳酪",
        "庫洛米造型蛋糕(芋頭戚風)",
        "3D庫洛米蛋糕(水果戚風)",
        "蝴蝶蕾絲邊庫洛米蛋糕(巧克力戚風)"
    ]
    return render_template('order.html', cakes=cakes)

@app.route('/success')
def success():
    """
    訂購成功頁面：顯示顧客剛剛填寫的訂購明細與提醒
    """
    order_data = session.get('last_order')
    email_info = session.get('email_info', '')
    
    # 如果 Session 中沒有訂單資料 (直接輸入網址)，就導回首頁
    if not order_data:
        return redirect(url_for('index'))
        
    return render_template('success.html', order=order_data, email_info=email_info)

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    """
    管理後台頁面：驗證登入後讀取 orders.xlsx，呈現所有訂單
    """
    if request.method == 'POST':
        password = request.form.get('password', '')
        if password == ADMIN_PASSWORD:
            session['admin_logged_in'] = True
            return redirect(url_for('admin'))
        else:
            return render_template('admin_login.html', error="密碼錯誤，請再試一次！ 😈")

    # 檢查是否已登入
    if not session.get('admin_logged_in'):
        return render_template('admin_login.html')

    all_orders = read_orders_from_excel()
    total_orders = len(all_orders)
    
    # 分頁設定 (一頁放十個訂單)
    page = request.args.get('page', 1, type=int)
    per_page = 10
    total_pages = (total_orders + per_page - 1) // per_page if total_orders > 0 else 1
    
    # 限制頁數在合理範圍內
    if page < 1:
        page = 1
    elif page > total_pages:
        page = total_pages
        
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    orders = all_orders[start_idx:end_idx]

    return render_template(
        'admin.html', 
        orders=orders, 
        page=page, 
        total_pages=total_pages, 
        total_orders=total_orders
    )

@app.route('/admin/logout')
def admin_logout():
    """
    登出管理後台
    """
    session.pop('admin_logged_in', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    # 啟動 Flask 伺服器，並開啟除錯模式以便開發
    app.run(debug=True, host='127.0.0.1', port=5000)
