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
        headers = ["序号", "用户名", "密码", "邮箱", "姓名", "年龄", "手机号", "注册状态", "登录状态", "购票姓名", "身份证号", "购票状态", "开始抢票时间", "点击抢票按钮次数", "抢票结果"]
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

def get_ticket_button_status(page):
    """
    获取抢票按钮的状态
    返回值: {
        'status': 'not_released'|'countdown'|'available'|'sold_out',
        'button': button_element,
        'text': button_text,
        'countdown_seconds': 倒计时秒数 (如果是countdown状态)
    }
    """
    # 查找抢票按钮
    button_selectors = [
        'button.btn-grab-ticket',
        'button:has-text("立即抢票")',
        'button:has-text("未放票")',
        'button:has-text("已售罄")',
        'button.ticket-btn',
        'button[class*="grab"]',
        'button[class*="ticket"]'
    ]
    
    ticket_btn = None
    for selector in button_selectors:
        try:
            ticket_btn = page.query_selector(selector)
            if ticket_btn:
                break
        except:
            continue
    
    # 如果没找到,尝试查找所有按钮
    if not ticket_btn:
        buttons = page.query_selector_all('button')
        for btn in buttons:
            btn_text = btn.text_content().strip() if btn.text_content() else ''
            if any(keyword in btn_text for keyword in ['抢票', '未放票', '已售罄', '倒计时']):
                ticket_btn = btn
                break
    
    if not ticket_btn:
        return {'status': 'not_found', 'button': None, 'text': '', 'countdown_seconds': 0}
    
    btn_text = ticket_btn.text_content().strip() if ticket_btn.text_content() else ''
    btn_color = ticket_btn.evaluate('el => window.getComputedStyle(el).backgroundColor') if ticket_btn else ''
    
    # 判断按钮状态
    if '未放票' in btn_text or '即将开售' in btn_text:
        return {'status': 'not_released', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': 0}
    
    elif '已售罄' in btn_text or '售完' in btn_text:
        return {'status': 'sold_out', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': 0}
    
    elif '立即抢票' in btn_text or '立即购买' in btn_text or 'btn-grab-ticket' in str(ticket_btn.get_attribute('class')):
        return {'status': 'available', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': 0}
    
    else:
        # 尝试解析倒计时
        # 倒计时格式可能是: "05:32" 或 "5分32秒" 或纯秒数
        import re
        countdown_match = re.search(r'(\d+):(\d+)', btn_text)
        if countdown_match:
            minutes = int(countdown_match.group(1))
            seconds = int(countdown_match.group(2))
            total_seconds = minutes * 60 + seconds
            return {'status': 'countdown', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': total_seconds}
        
        # 尝试匹配 "X分Y秒" 格式
        countdown_match2 = re.search(r'(\d+)\s*分\s*(\d+)\s*秒', btn_text)
        if countdown_match2:
            minutes = int(countdown_match2.group(1))
            seconds = int(countdown_match2.group(2))
            total_seconds = minutes * 60 + seconds
            return {'status': 'countdown', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': total_seconds}
        
        # 如果是数字(纯秒数)
        if btn_text.isdigit():
            return {'status': 'countdown', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': int(btn_text)}
    
    # 默认返回可用状态
    return {'status': 'available', 'button': ticket_btn, 'text': btn_text, 'countdown_seconds': 0}


def wait_and_click_countdown(page, btn_info, task_id, start_time_before=10):
    """
    等待倒计时,在倒计时结束前start_time_before秒开始点击
    每秒点击1次
    返回: (click_count, success)
    """
    countdown_seconds = btn_info['countdown_seconds']
    button = btn_info['button']
    
    print(f"[任务{task_id}] 检测到倒计时: {countdown_seconds}秒")
    
    # 计算需要等待的时间
    wait_time = max(0, countdown_seconds - start_time_before)
    
    if wait_time > 0:
        print(f"[任务{task_id}] 等待 {wait_time} 秒后开始抢票...")
        time.sleep(wait_time)
    
    # 开始抢票,每秒点击1次
    click_count = 0
    start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    print(f"[任务{task_id}] 开始抢票点击 (提前{start_time_before}秒), 每秒1次...")
    
    # 持续点击直到成功或倒计时结束后再点几次
    max_attempts = start_time_before + 5  # 多尝试5秒
    
    for attempt in range(max_attempts):
        try:
            # 重新获取按钮状态
            current_status = get_ticket_button_status(page)
            
            if current_status['status'] == 'available':
                # 按钮变为可点击状态
                print(f"[任务{task_id}] 按钮变为可点击状态,立即点击!")
                current_status['button'].click()
                click_count += 1
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
                return click_count, True, start_time
            
            elif current_status['status'] == 'sold_out':
                print(f"[任务{task_id}] 票已售罄!")
                return click_count, False, start_time
            
            elif current_status['status'] == 'countdown' and current_status['button']:
                # 继续点击
                current_status['button'].click()
                click_count += 1
                print(f"[任务{task_id}] 第 {click_count} 次点击抢票按钮")
                time.sleep(1)
            
            elif current_status['status'] == 'not_found':
                # 页面可能已跳转
                print(f"[任务{task_id}] 按钮状态变化,可能已跳转")
                return click_count, True, start_time
            
            else:
                # 其他情况也尝试点击
                if current_status['button']:
                    current_status['button'].click()
                    click_count += 1
                    print(f"[任务{task_id}] 第 {click_count} 次点击抢票按钮")
                time.sleep(1)
                
        except Exception as e:
            print(f"[任务{task_id}] 点击时出错: {e}")
            # 可能是页面已跳转,视为成功
            return click_count, True, start_time
    
    return click_count, False, start_time


def check_ticket_release(page, task_id, check_interval=180):
    """
    检查是否放票,每check_interval秒检查一次
    返回: button_info 当按钮状态变为非未放票时
    """
    print(f"[任务{task_id}] 票未放出,开始轮询检查 (每{check_interval}秒检查一次)...")
    
    check_count = 0
    while True:
        check_count += 1
        print(f"[任务{task_id}] 第 {check_count} 次检查票状态...")
        
        btn_info = get_ticket_button_status(page)
        
        if btn_info['status'] != 'not_released':
            print(f"[任务{task_id}] 票状态变化: {btn_info['status']}")
            return btn_info
        
        print(f"[任务{task_id}] 仍未放票, {check_interval}秒后再次检查...")
        time.sleep(check_interval)
        
        # 刷新页面获取最新状态
        try:
            page.reload()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
        except Exception as e:
            print(f"[任务{task_id}] 刷新页面出错: {e}")


def perform_ticket_purchase(page, context, task_id):
    """
    新的抢票流程
    1. 判断抢票按钮状态
    2. 根据状态执行不同策略
    3. 记录抢票数据
    """
    print(f"[任务{task_id}] 开始抢票流程...")
    
    # 首先进入抢票页面
    ticket_btn = page.query_selector('button.btn-grab-ticket') or page.query_selector('button:has-text("立即抢票")')
    if ticket_btn:
        print(f"[任务{task_id}] 点击立即抢票按钮进入抢票页面...")
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
    
    # 获取抢票页面
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
            "start_time": "",
            "click_count": 0,
            "result": "失败"
        }
    
    ticket_page.wait_for_load_state("networkidle")
    ticket_page.wait_for_timeout(2000)
    print(f"[任务{task_id}] 抢票页面URL: {ticket_page.url}")
    
    # 初始化抢票记录数据
    start_time = ""
    click_count = 0
    result = "失败"
    
    # 判断抢票按钮状态并执行相应策略
    btn_info = get_ticket_button_status(ticket_page)
    print(f"[任务{task_id}] 抢票按钮状态: {btn_info['status']}, 文字: {btn_info['text']}")
    
    if btn_info['status'] == 'not_released':
        # 未放票状态,每3分钟检查一次
        btn_info = check_ticket_release(ticket_page, task_id, check_interval=180)
        # 检查后继续判断新状态
    
    if btn_info['status'] == 'sold_out':
        # 已售罄
        print(f"[任务{task_id}] 票已售罄,抢票失败!")
        result = "失败-已售罄"
        return {
            "ticket_name": "",
            "ticket_id": "",
            "purchase_status": "失败-已售罄",
            "start_time": "",
            "click_count": 0,
            "result": "失败-已售罄"
        }
    
    if btn_info['status'] == 'countdown':
        # 倒计时状态,提前10秒开始点击,每秒1次
        click_count, success, start_time = wait_and_click_countdown(ticket_page, btn_info, task_id, start_time_before=10)
        if not success:
            return {
                "ticket_name": "",
                "ticket_id": "",
                "purchase_status": "失败-倒计时抢票未成功",
                "start_time": start_time,
                "click_count": click_count,
                "result": "失败"
            }
        # 成功后继续填写信息
    
    elif btn_info['status'] == 'available':
        # 立即抢票状态
        start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[任务{task_id}] 立即抢票,开始时间: {start_time}")
        
        try:
            btn_info['button'].click()
            click_count = 1
            ticket_page.wait_for_load_state("networkidle")
            ticket_page.wait_for_timeout(3000)
        except Exception as e:
            print(f"[任务{task_id}] 点击抢票按钮出错: {e}")
    
    # 填写购票信息
    print(f"[任务{task_id}] 购票页面URL: {ticket_page.url}")
    
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
            result = "成功"
        else:
            purchase_status = "失败-未找到提交按钮"
            result = "失败"
    else:
        purchase_status = "失败-未找到输入框"
        result = "失败"
    
    return {
        "ticket_name": ticket_name,
        "ticket_id": ticket_id,
        "purchase_status": purchase_status,
        "start_time": start_time,
        "click_count": click_count,
        "result": result
    }

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
                "start_time": "",
                "click_count": 0,
                "result": "未尝试"
            }

            if login_result['login_status'] == "成功":
                ticket_result = perform_ticket_purchase(page, context, task_id)
                print(f"[任务{task_id}] 购票状态: {ticket_result['purchase_status']}")

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
                ticket_result.get("start_time", ""),
                ticket_result.get("click_count", 0),
                ticket_result.get("result", "")
            ]

            save_to_excel(excel_data)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 注册状态: {register_result['register_status']}")
            print(f"[任务{task_id}] 登录状态: {login_result['login_status']}")
            print(f"[任务{task_id}] 购票状态: {ticket_result['purchase_status']}")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "register_status": register_result["register_status"],
                "login_status": login_result["login_status"],
                "purchase_status": ticket_result["purchase_status"]
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
