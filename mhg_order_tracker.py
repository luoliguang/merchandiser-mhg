"""
棉花果(58mhg.com) 订单进度自动查询爬虫
依赖: pip install playwright pandas openpyxl
初次运行: playwright install chromium
"""

import asyncio
import json
import time
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright, Page


# ========== 配置区 ==========
CONFIG = {
    "url": "https://58mhg.com/console/#/workbench",
    "login_url": "https://58mhg.com/console/#/login",
    "username": "17783386425",       # ← 改成你的账号(手机号/邮箱)
    "password": "zxcvbnm123..",       # ← 改成你的密码
    "headless": False,            # False = 显示浏览器窗口，便于调试；True = 无头模式
    "output_file": f"orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    "query_interval_seconds": 60, # 定时轮询间隔（秒），0 = 只查一次
}
# ============================


async def login(page: Page):
    """登录棉花果"""
    print("正在跳转登录页...")
    await page.goto(CONFIG["login_url"], wait_until="networkidle")
    await page.wait_for_timeout(2000)

    # 输入账号密码（选择器可能需要根据实际页面调整）
    try:
        # 尝试常见的输入框选择器
        await page.fill('input[type="text"], input[placeholder*="账号"], input[placeholder*="手机"], input[name="username"]',
                        CONFIG["username"])
        await page.fill('input[type="password"]', CONFIG["password"])
        
        # 点击登录按钮
        await page.click('button[type="submit"], button:has-text("登录")')
        await page.wait_for_timeout(3000)
        print("✅ 登录成功")
    except Exception as e:
        print(f"❌ 自动登录失败，请手动在浏览器中登录: {e}")
        print("等待30秒，请手动完成登录...")
        await page.wait_for_timeout(30000)


async def fetch_orders_via_api(page: Page) -> list:
    """
    方法1: 拦截网络请求获取订单数据（推荐，速度快）
    """
    orders = []
    
    async def handle_response(response):
        """拦截包含订单数据的API响应"""
        url = response.url
        # 根据实际API路径调整关键词
        if any(kw in url for kw in ["order", "orders", "订单"]):
            try:
                data = await response.json()
                print(f"  捕获API: {url}")
                # 尝试提取订单列表（根据实际数据结构调整）
                if isinstance(data, dict):
                    records = data.get("data", data.get("list", data.get("records", [])))
                    if isinstance(records, list):
                        orders.extend(records)
                elif isinstance(data, list):
                    orders.extend(data)
            except:
                pass
    
    page.on("response", handle_response)
    
    # 导航到订单页面
    print("正在跳转订单页面...")
    await page.goto("https://58mhg.com/console/#/order", wait_until="networkidle")
    await page.wait_for_timeout(3000)
    
    # 如果没有捕获到订单，尝试点击"订单管理"菜单
    if not orders:
        try:
            order_menu = page.locator('text=订单管理, text=我的订单, text=订单列表').first
            await order_menu.click()
            await page.wait_for_timeout(3000)
        except:
            pass
    
    return orders


async def fetch_orders_via_scraping(page: Page) -> list:
    """
    方法2: 直接解析页面DOM获取订单数据（备用）
    """
    orders = []
    
    print("正在解析订单页面...")
    await page.wait_for_timeout(2000)
    
    # 尝试查找订单表格行
    rows = await page.query_selector_all("table tbody tr, .order-item, .order-row")
    
    for row in rows:
        try:
            text = await row.inner_text()
            cells = await row.query_selector_all("td")
            row_data = [await cell.inner_text() for cell in cells]
            if row_data:
                orders.append({"raw": " | ".join(row_data)})
        except:
            continue
    
    return orders


async def save_orders(orders: list, filename: str):
    """保存订单数据到Excel"""
    if not orders:
        print("⚠️ 未获取到订单数据")
        return
    
    df = pd.DataFrame(orders)
    df.to_excel(filename, index=False)
    print(f"✅ 已保存 {len(orders)} 条订单到 {filename}")
    return df


async def print_order_summary(orders: list):
    """打印订单摘要"""
    if not orders:
        print("暂无订单数据")
        return
    
    print(f"\n{'='*50}")
    print(f"共找到 {len(orders)} 条订单记录")
    print(f"查询时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}")
    
    # 打印前10条
    for i, order in enumerate(orders[:10], 1):
        if "raw" in order:
            print(f"{i}. {order['raw']}")
        else:
            # 尝试打印关键字段
            status = order.get("status", order.get("orderStatus", order.get("state", "未知")))
            order_no = order.get("orderNo", order.get("orderId", order.get("id", "未知")))
            print(f"{i}. 订单号: {order_no} | 状态: {status}")
    
    if len(orders) > 10:
        print(f"... 还有 {len(orders)-10} 条，已保存到文件")


async def main():
    print("🚀 棉花果订单查询爬虫启动")
    print(f"目标网站: {CONFIG['url']}")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=CONFIG["headless"])
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = await context.new_page()
        
        # 1. 登录
        await login(page)
        
        # 2. 查询订单（循环）
        run_count = 0
        while True:
            run_count += 1
            print(f"\n--- 第 {run_count} 次查询 ---")
            
            try:
                # 优先使用API拦截方式
                orders = await fetch_orders_via_api(page)
                
                # 如果API方式没有获取到，使用DOM解析
                if not orders:
                    orders = await fetch_orders_via_scraping(page)
                
                # 打印摘要
                await print_order_summary(orders)
                
                # 保存到文件
                if orders:
                    filename = CONFIG["output_file"] if run_count == 1 else \
                               f"orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    await save_orders(orders, filename)
                    
            except Exception as e:
                print(f"❌ 查询出错: {e}")
            
            # 是否循环轮询
            if CONFIG["query_interval_seconds"] <= 0:
                break
            
            print(f"\n⏳ {CONFIG['query_interval_seconds']} 秒后进行下一次查询... (Ctrl+C 退出)")
            await asyncio.sleep(CONFIG["query_interval_seconds"])
        
        await browser.close()
        print("\n✅ 爬虫运行完毕")


if __name__ == "__main__":
    asyncio.run(main())
