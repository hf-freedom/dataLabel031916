import random
import string
import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "register_data.xlsx")

FIRST_NAMES = ["张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗"]
LAST_NAMES = ["伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军", "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞"]

excel_lock = threading.Lock()

def generate_random_string(length=8):
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=length))

def generate_random_email():
    username = generate_random_string(10)
    return f"{username}@163.com"

def generate_random_password():
    return generate_random_string(12) + random.choice(string.ascii_uppercase) + random.choice(string.digits)

def generate_random_name():
    first_name = random.choice(FIRST_NAMES)
    last_name = random.choice(LAST_NAMES)
    return first_name + last_name

def generate_random_age():
    return str(random.randint(18, 60))

def generate_random_phone():
    prefixes = ["130", "131", "132", "133", "134", "135", "136", "137", "138", "139",
                "150", "151", "152", "153", "155", "156", "157", "158", "159",
                "170", "176", "177", "178",
                "180", "181", "182", "183", "184", "185", "186", "187", "188", "189"]
    prefix = random.choice(prefixes)
    suffix = ''.join(random.choices(string.digits, k=8))
    return prefix + suffix

def generate_id_card():
    area_code = random.choice(["110101", "310101", "440101", "440301", "330101", "320101", "510101", "420101"])
    year = random.randint(1970, 2000)
    month = random.randint(1, 12)
    day = random.randint(1, 28)
    serial = random.randint(100, 999)
    
    id_str = f"{area_code}{year:04d}{month:02d}{day:02d}{serial:03d}"
    
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    check_codes = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
    
    total = sum(int(id_str[i]) * weights[i] for i in range(17))
    check_code = check_codes[total % 11]
    
    return id_str + check_code

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "注册购票数据"
        headers = ["序号", "用户名", "密码", "邮箱", "姓名", "年龄", "手机号", "注册状态", "登录状态", "购票姓名", "身份证号", "购票状态", "开始抢票时间", "点击抢票按钮次数", "是否成功"]
        ws.append(headers)
        wb.save(EXCEL_FILE)
        print(f"创建Excel文件: {EXCEL_FILE}")
    return EXCEL_FILE

def save_to_excel(data):
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)

def find_input(page, selectors, field_name):
    for selector in selectors:
        try:
            element = page.query_selector(selector)
            if element:
                return element, selector
        except:
            continue
    return None, None

def check_ticket_button_status(page, task_id):
    button_selectors = [
        'button.btn-grab-ticket',
        'button.ticket-btn',
        'button:has-text("抢票")',
        'button:has-text("立即抢票")',
        'button:has-text("立即购票")',
        'button:has-text("未放票")',
        'button:has-text("已售罄")',
        'button'
    ]
    
    for selector in button_selectors:
        try:
            btn = page.query_selector(selector)
            if btn:
                btn_text = btn.text_content().strip() if btn.text_content() else ''
                btn_class = btn.get_attribute('class') or ''
                btn_style = btn.get_attribute('style') or ''
                is_disabled = btn.is_disabled()
                
                print(f"[任务{task_id}] 按钮状态检测 - 文本: '{btn_text}', Class: '{btn_class}', 禁用: {is_disabled}")
                
                if '未放票' in btn_text:
                    return {'status': 'not_available', 'button': btn, 'text': btn_text}
                elif '已售罄' in btn_text or '售罄' in btn_text:
                    return {'status': 'sold_out', 'button': btn, 'text': btn_text}
                elif '立即抢票' in btn_text or '立即购票' in btn_text or '抢票' in btn_text:
                    return {'status': 'available', 'button': btn, 'text': btn_text}
                elif any(c.isdigit() for c in btn_text) and (':' in btn_text or '倒计时' in btn_text):
                    return {'status': 'countdown', 'button': btn, 'text': btn_text}
        except:
            continue
    
    buttons = page.query_selector_all('button')
    for btn in buttons:
        try:
            btn_text = btn.text_content().strip() if btn.text_content() else ''
            if btn_text:
                print(f"[任务{task_id}] 发现按钮 - 文本: '{btn_text}'")
                if '未放票' in btn_text:
                    return {'status': 'not_available', 'button': btn, 'text': btn_text}
                elif '已售罄' in btn_text or '售罄' in btn_text:
                    return {'status': 'sold_out', 'button': btn, 'text': btn_text}
                elif '抢票' in btn_text or '购票' in btn_text:
                    return {'status': 'available', 'button': btn, 'text': btn_text}
                elif any(c.isdigit() for c in btn_text) and (':' in btn_text or '倒计时' in btn_text):
                    return {'status': 'countdown', 'button': btn, 'text': btn_text}
        except:
            continue
    
    return {'status': 'not_found', 'button': None, 'text': ''}

def parse_countdown(btn_text):
    try:
        digits = ''.join([c for c in btn_text if c.isdigit() or c == ':'])
        if ':' in digits:
            parts = digits.split(':')
            if len(parts) == 2:
                minutes = int(parts[0])
                seconds = int(parts[1])
                return minutes * 60 + seconds
            elif len(parts) == 3:
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = int(parts[2])
                return hours * 3600 + minutes * 60 + seconds
        return None
    except:
        return None

def perform_register(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    email = user_data["email"]
    name = user_data["name"]
    age = user_data["age"]
    phone = user_data["phone"]
    
    username_selectors = [
        'input[name="username"]',
        'input[name="user"]',
        'input[name="userName"]',
        'input[placeholder*="用户名"]',
        'input[placeholder*="账号"]',
        '#username',
        '#userName',
        'input[type="text"]:first-of-type'
    ]
    
    password_selectors = [
        'input[name="password"]',
        'input[name="pwd"]',
        'input[name="userPassword"]',
        'input[placeholder*="密码"]',
        '#password',
        '#pwd',
        'input[type="password"]'
    ]
    
    email_selectors = [
        'input[name="email"]',
        'input[name="mail"]',
        'input[placeholder*="邮箱"]',
        'input[placeholder*="Email"]',
        '#email',
        'input[type="email"]'
    ]
    
    name_selectors = [
        'input[name="name"]',
        'input[name="realName"]',
        'input[name="realname"]',
        'input[name="userName"]',
        'input[placeholder*="姓名"]',
        'input[placeholder*="真实姓名"]',
        '#name',
        '#realName',
        '#userName'
    ]
    
    age_selectors = [
        'input[name="age"]',
        'input[placeholder*="年龄"]',
        '#age',
        'input[type="number"]'
    ]
    
    phone_selectors = [
        'input[name="phone"]',
        'input[name="mobile"]',
        'input[name="tel"]',
        'input[name="phoneNumber"]',
        'input[placeholder*="手机"]',
        'input[placeholder*="电话"]',
        'input[placeholder*="手机号"]',
        '#phone',
        '#mobile',
        '#tel'
    ]
    
    print(f"[任务{task_id}] 查找注册表单字段...")
    
    input_elements = page.query_selector_all('input')
    if len(input_elements) >= 4:
        input_elements[0].fill(username)  # 账号
        input_elements[1].fill(name)      # 姓名
        input_elements[3].fill(password)  # 密码
    
    page.wait_for_timeout(500)
    
    submit_selectors = [
        'button:has-text("注册")',
        'button:has-text("提交")',
        'button:has-text("确定")',
        'input[type="submit"]',
        'input[value="注册"]',
        'input[value="提交"]',
        '.register-btn',
        '#register-btn',
        'button[type="submit"]'
    ]
    
    submit_btn = None
    for selector in submit_selectors:
        try:
            submit_btn = page.query_selector(selector)
            if submit_btn:
                break
        except:
            continue
    
    register_status = "失败"
    if submit_btn:
        print(f"[任务{task_id}] 点击注册按钮...")
        submit_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        register_status = "成功"
    
    return {
        "register_status": register_status
    }

def perform_login(page, user_data, task_id, context):
    username = user_data["username"]
    password = user_data["password"]
    
    print(f"[任务{task_id}] 跳转到登录页面...")
    
    if 'register-success' in page.url:
        login_now_btn = page.query_selector('button.btn-primary') or page.query_selector('button:has-text("立即登录")')
        if login_now_btn:
            login_now_btn.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
    else:
        login_link_selectors = [
            'a:has-text("登录")',
            'text=登录',
            'a[href*="login"]',
            '.login-link',
            '#login-link'
        ]
        
        login_link = None
        for selector in login_link_selectors:
            try:
                login_link = page.query_selector(selector)
                if login_link:
                    break
            except:
                continue
        
        if login_link:
            login_link.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
        else:
            page.goto("http://39.107.109.8:8082/login", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
    
    login_inputs = page.query_selector_all('input')
    if len(login_inputs) >= 2:
        login_inputs[0].fill(username)
        login_inputs[1].fill(password)
    
    page.wait_for_timeout(500)
    
    login_btn_selectors = [
        'button:has-text("登录")',
        'button:has-text("Login")',
        'input[type="submit"]',
        'input[value="登录"]',
        '.login-btn',
        '#login-btn',
        'button[type="submit"]'
    ]
    
    login_btn = None
    for selector in login_btn_selectors:
        try:
            login_btn = page.query_selector(selector)
            if login_btn:
                break
        except:
            continue
    
    login_status = "失败"
    
    if login_btn:
        print(f"[任务{task_id}] 点击登录按钮...")
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        login_status = "成功"
    
    return {
        "login_status": login_status
    }

def fill_ticket_info(ticket_page, task_id):
    inputs = ticket_page.query_selector_all('input')
    print(f"[任务{task_id}] 购票表单输入框数量: {len(inputs)}")
    
    ticket_name = generate_random_name()
    ticket_id = generate_id_card()
    
    print(f"[任务{task_id}] 生成购票信息 - 姓名: {ticket_name}, 身份证: {ticket_id}")
    
    purchase_status = "失败"
    
    if len(inputs) >= 2:
        for inp in inputs:
            placeholder = inp.get_attribute('placeholder') or ''
            if '姓名' in placeholder or 'name' in placeholder.lower():
                inp.fill(ticket_name)
            elif '身份证' in placeholder or 'id' in placeholder.lower() or 'card' in placeholder.lower():
                inp.fill(ticket_id)
        
        ticket_page.wait_for_timeout(1000)
        
        submit_btn = None
        buttons = ticket_page.query_selector_all('button')
        for btn in buttons:
            btn_text = btn.text_content().strip() if btn.text_content() else ''
            if '确认' in btn_text or '购买' in btn_text or '提交' in btn_text:
                submit_btn = btn
                break
        
        if submit_btn:
            print(f"[任务{task_id}] 点击确认购买按钮...")
            submit_btn.click()
            ticket_page.wait_for_load_state("networkidle")
            ticket_page.wait_for_timeout(2000)
            purchase_status = "成功"
        else:
            purchase_status = "失败-未找到提交按钮"
    else:
        purchase_status = "失败-未找到输入框"
    
    return {
        "ticket_name": ticket_name,
        "ticket_id": ticket_id,
        "purchase_status": purchase_status
    }

def perform_ticket_purchase(page, context, task_id):
    print(f"[任务{task_id}] 开始抢票流程...")
    
    start_ticket_time = ""
    click_count = 0
    is_success = "否"
    
    ticket_btn = page.query_selector('button.btn-grab-ticket') or page.query_selector('button:has-text("立即抢票")')
    if ticket_btn:
        print(f"[任务{task_id}] 点击立即抢票按钮...")
        ticket_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
    else:
        buttons = page.query_selector_all('button')
        for btn in buttons:
            btn_text = btn.text_content().strip() if btn.text_content() else ''
            if '抢票' in btn_text:
                btn.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(3000)
                break
    
    ticket_page = None
    for p in context.pages:
        if '8085' in p.url or 'ticket' in p.url.lower():
            ticket_page = p
            break
    
    if not ticket_page and len(context.pages) > 1:
        ticket_page = context.pages[1]
    
    if not ticket_page:
        print(f"[任务{task_id}] 未找到抢票页面")
        return {
            "ticket_name": "",
            "ticket_id": "",
            "purchase_status": "失败-未找到抢票页面",
            "start_ticket_time": start_ticket_time,
            "click_count": click_count,
            "is_success": is_success
        }
    
    ticket_page.wait_for_load_state("networkidle")
    ticket_page.wait_for_timeout(2000)
    print(f"[任务{task_id}] 抢票页面URL: {ticket_page.url}")
    
    while True:
        btn_status = check_ticket_button_status(ticket_page, task_id)
        current_status = btn_status['status']
        
        if current_status == 'not_available':
            print(f"[任务{task_id}] 按钮状态: 未放票，3分钟后再次检查...")
            ticket_page.wait_for_timeout(180000)
            ticket_page.reload()
            ticket_page.wait_for_load_state("networkidle")
            continue
            
        elif current_status == 'sold_out':
            print(f"[任务{task_id}] 按钮状态: 已售罄，抢票失败")
            return {
                "ticket_name": "",
                "ticket_id": "",
                "purchase_status": "失败-已售罄",
                "start_ticket_time": start_ticket_time,
                "click_count": click_count,
                "is_success": is_success
            }
            
        elif current_status == 'countdown':
            print(f"[任务{task_id}] 按钮状态: 倒计时 - {btn_status['text']}")
            countdown_seconds = parse_countdown(btn_status['text'])
            
            if countdown_seconds and countdown_seconds > 10:
                wait_seconds = countdown_seconds - 10
                print(f"[任务{task_id}] 倒计时解析: {countdown_seconds}秒，将在{wait_seconds}秒后开始点击...")
                ticket_page.wait_for_timeout(wait_seconds * 1000)
            
            start_ticket_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[任务{task_id}] 开始抢票时间: {start_ticket_time}")
            
            for i in range(20):
                try:
                    ticket_page.wait_for_timeout(500)
                    current_btn_status = check_ticket_button_status(ticket_page, task_id)
                    if current_btn_status['status'] == 'available':
                        btn = current_btn_status['button']
                        if btn and not btn.is_disabled():
                            btn.click()
                            click_count += 1
                            print(f"[任务{task_id}] 第{click_count}次点击抢票按钮")
                            ticket_page.wait_for_timeout(500)
                    elif current_btn_status['status'] == 'countdown':
                        continue
                    else:
                        break
                except Exception as e:
                    print(f"[任务{task_id}] 点击时发生错误: {e}")
                    ticket_page.wait_for_timeout(500)
                    continue
            
            ticket_page.wait_for_load_state("networkidle")
            ticket_page.wait_for_timeout(2000)
            break
            
        elif current_status == 'available':
            print(f"[任务{task_id}] 按钮状态: 立即抢票，开始购票流程")
            start_ticket_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            btn = btn_status['button']
            if btn:
                btn.click()
                click_count += 1
                print(f"[任务{task_id}] 点击抢票按钮，第{click_count}次")
                ticket_page.wait_for_load_state("networkidle")
                ticket_page.wait_for_timeout(2000)
            break
            
        else:
            print(f"[任务{task_id}] 未找到明确状态的按钮，继续等待...")
            ticket_page.wait_for_timeout(5000)
            continue
    
    print(f"[任务{task_id}] 购票页面URL: {ticket_page.url}")
    result = fill_ticket_info(ticket_page, task_id)
    
    if result['purchase_status'] == "成功":
        is_success = "是"
    
    result['start_ticket_time'] = start_ticket_time
    result['click_count'] = click_count
    result['is_success'] = is_success
    
    return result

def single_task(task_id, user_data):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print(f"[任务{task_id}] 正在打开网站...")
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            print(f"[任务{task_id}] 查找注册链接...")
            register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
            
            if register_link:
                print(f"[任务{task_id}] 点击注册链接...")
                register_link.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
            
            register_result = perform_register(page, user_data, task_id)
            print(f"[任务{task_id}] 注册状态: {register_result['register_status']}")
            
            login_result = perform_login(page, user_data, task_id, context)
            print(f"[任务{task_id}] 登录状态: {login_result['login_status']}")
            
            ticket_result = {
                "ticket_name": "",
                "ticket_id": "",
                "purchase_status": "未购票",
                "start_ticket_time": "",
                "click_count": 0,
                "is_success": "否"
            }
            
            if login_result['login_status'] == "成功":
                ticket_result = perform_ticket_purchase(page, context, task_id)
                print(f"[任务{task_id}] 购票状态: {ticket_result['purchase_status']}")
                print(f"[任务{task_id}] 开始抢票时间: {ticket_result['start_ticket_time']}")
                print(f"[任务{task_id}] 点击抢票按钮次数: {ticket_result['click_count']}")
                print(f"[任务{task_id}] 是否成功: {ticket_result['is_success']}")
            
            excel_data = [
                task_id,
                user_data["username"],
                user_data["password"],
                user_data["email"],
                user_data["name"],
                user_data["age"],
                user_data["phone"],
                register_result["register_status"],
                login_result["login_status"],
                ticket_result["ticket_name"],
                ticket_result["ticket_id"],
                ticket_result["purchase_status"],
                ticket_result["start_ticket_time"],
                ticket_result["click_count"],
                ticket_result["is_success"]
            ]
            
            save_to_excel(excel_data)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 注册状态: {register_result['register_status']}")
            print(f"[任务{task_id}] 登录状态: {login_result['login_status']}")
            print(f"[任务{task_id}] 购票状态: {ticket_result['purchase_status']}")
            print(f"[任务{task_id}] 开始抢票时间: {ticket_result['start_ticket_time']}")
            print(f"[任务{task_id}] 点击抢票按钮次数: {ticket_result['click_count']}")
            print(f"[任务{task_id}] 是否成功: {ticket_result['is_success']}")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "register_status": register_result["register_status"],
                "login_status": login_result["login_status"],
                "purchase_status": ticket_result["purchase_status"],
                "start_ticket_time": ticket_result["start_ticket_time"],
                "click_count": ticket_result["click_count"],
                "is_success": ticket_result["is_success"]
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
            
            return {
                "task_id": task_id,
                "status": "失败",
                "error": str(e)
            }
        finally:
            browser.close()

def generate_user_data():
    return {
        "username": "user_" + generate_random_string(6),
        "password": generate_random_password(),
        "email": generate_random_email(),
        "name": generate_random_name(),
        "age": generate_random_age(),
        "phone": generate_random_phone()
    }

def run_parallel_register(num_tasks=5):
    init_excel()
    
    print("=" * 60)
    print(f"开始并行执行 {num_tasks} 个注册购票任务")
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print(f"  任务{i}: {user['username']}")
    
    results = []
    
    with ThreadPoolExecutor(max_workers=num_tasks) as executor:
        futures = {executor.submit(single_task, i+1, user): i+1 for i, user in enumerate(users_data)}
        
        for future in as_completed(futures):
            task_id = futures[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                print(f"任务{task_id}执行异常: {e}")
                results.append({"task_id": task_id, "status": "异常", "error": str(e)})
    
    overall_end_time = datetime.now()
    overall_duration = (overall_end_time - overall_start_time).total_seconds()
    
    print("\n" + "=" * 60)
    print("所有任务执行完成!")
    print("=" * 60)
    
    success_count = sum(1 for r in results if r.get("status") == "成功")
    purchase_success_count = sum(1 for r in results if r.get("purchase_status") == "成功")
    fail_count = num_tasks - success_count
    
    print(f"\n执行统计:")
    print(f"  总任务数: {num_tasks}")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    print(f"  购票成功: {purchase_success_count}")
    print(f"  总耗时: {overall_duration:.2f}秒")
    print(f"  平均耗时: {overall_duration/num_tasks:.2f}秒/任务")
    
    print(f"\n数据已保存到: {EXCEL_FILE}")
    
    return results

if __name__ == "__main__":
    results = run_parallel_register(2)
