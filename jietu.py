# jietu.py
"""
截图模块 - 使用Selenium + Edge浏览器截取网页viewport（可视区域）
只截取网页显示内容区域，不包括浏览器标签页和工具栏
确保截取到控制面板滚动后的最新状态
"""

import os
import time
import socket
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException


def capture_screenshot(url: str, save_dir: str = "网点图", wait_time: int = 2, ui_state: dict = None, driver_instance = None, group_name: str = "") -> str:
    """
    使用Selenium + Edge浏览器截图页面viewport（可视区域）
    确保截取到控制面板滚动后的最新状态，并同步当前浏览器中的UI状态
    
    Args:
        url: 要截图的网页URL
        save_dir: 保存目录，默认为"网点图"
        wait_time: 等待页面加载和滚动完成的时间（秒），默认2秒
        ui_state: UI状态字典，包含 showFarthestLine 和 showDistanceLabels 等状态
        driver_instance: 浏览器实例（可选）
        group_name: 网组名称，用于截图文件命名（可选）
    
    Returns:
        保存的文件路径
    
    Raises:
        Exception: 截图失败时抛出异常
    """
    if ui_state is None:
        ui_state = {}
    # 创建保存目录
    os.makedirs(save_dir, exist_ok=True)
    
    # 生成文件名（优先使用网组名称，否则使用时间戳）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if group_name and group_name.strip():
        # 清理网组名称，移除非法文件名字符
        safe_group_name = "".join(c for c in group_name.strip() if c.isalnum() or c in ('-', '_', ' ')).strip()
        safe_group_name = safe_group_name.replace(' ', '_')
        filename = f"{safe_group_name}_{timestamp}.png"
    else:
        filename = f"route_screenshot_{timestamp}.png"
    filepath = os.path.join(save_dir, filename)
    
    driver = None
    should_close_driver = False  # 标记是否需要关闭driver（如果是外部传入的，不关闭）
    
    try:
        print(f"[截图] 开始截图: {url}")
        
        # 必须使用app.py传入的浏览器实例，不启动新浏览器
        if driver_instance is None:
            error_msg = "未传入浏览器实例。请确保app.py已启动并打开了浏览器，截图功能将使用该浏览器实例。"
            print(f"[截图] ❌ {error_msg}")
            raise Exception(error_msg)
        
        # 检查浏览器实例是否仍然有效（包括窗口是否打开）
        try:
            # 首先检查窗口句柄是否存在（如果窗口被关闭，这个会失败）
            window_handles = driver_instance.window_handles
            if not window_handles:
                raise Exception("浏览器窗口已被关闭（没有活动窗口）")
            
            # 切换到第一个窗口（确保有活动窗口）
            driver_instance.switch_to.window(window_handles[0])
            
            # 尝试获取当前URL来验证浏览器是否仍然有效
            current_url = driver_instance.current_url
            print(f"[截图] ✓ 使用 app.py 中已打开的浏览器实例")
            print(f"[截图]   当前浏览器URL: {current_url}")
            print(f"[截图]   活动窗口数量: {len(window_handles)}")
            
            # 检查页面是否已经加载了应用（通过检查关键元素）
            # 如果页面已经加载，就不刷新，避免丢失已渲染的内容（如路线连线等）
            try:
                # 检查控制面板是否存在（说明应用已加载）
                from selenium.webdriver.common.by import By
                control_panel = driver_instance.find_elements(By.ID, "control-panel")
                # 如果当前URL包含目标URL的基础部分，且控制面板存在，说明页面已加载
                if control_panel and (url in current_url or current_url.startswith(url.split('?')[0].split('#')[0])):
                    print(f"[截图] ✓ 页面已加载应用，无需刷新（避免丢失已渲染内容）")
                else:
                    # 只有在页面确实不在应用页面时，才刷新
                    print(f"[截图] 页面未加载应用，正在导航到: {url}")
                    driver_instance.get(url)
                    # 等待页面加载
                    time.sleep(2)
            except Exception as check_e:
                # 如果检查失败，尝试导航到目标URL
                print(f"[截图] 检查页面状态失败: {check_e}，尝试导航到: {url}")
                driver_instance.get(url)
                # 等待页面加载
                time.sleep(2)
            
            driver = driver_instance
            should_close_driver = False  # 不关闭外部传入的driver
        except Exception as e:
            error_msg = f"浏览器实例无效（可能窗口已被关闭）: {e}。请重新启动app.py或确保浏览器窗口保持打开。"
            print(f"[截图] ❌ {error_msg}")
            raise Exception(error_msg)
        
        # 等待页面基本元素加载完成
        print("[截图] 等待页面元素加载...")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "control-panel"))
            )
            print("[截图] 控制面板元素已加载")
        except TimeoutException:
            print("[截图] ⚠️ 控制面板元素加载超时，继续...")
        
        # 等待页面内容渲染
        print(f"[截图] 等待 {wait_time} 秒以确保内容完全渲染...")
        time.sleep(wait_time)
        
        # 执行JavaScript：滚动控制面板到底部（确保显示最新内容）
        print("[截图] 滚动控制面板到底部...")
        try:
            scroll_script = """
            const panel = document.getElementById('control-panel');
            if (panel) {
                // 立即滚动到底部（不使用smooth，确保快速完成）
                panel.scrollTop = panel.scrollHeight;
                return true;
            }
            return false;
            """
            result = driver.execute_script(scroll_script)
            if result:
                print("[截图] ✓ 控制面板已滚动到底部")
            else:
                print("[截图] ⚠️ 未找到控制面板元素")
            
            # 等待一下，确保滚动和内容渲染完成
            time.sleep(0.8)
        except Exception as e:
            print(f"[截图] ⚠️ 滚动控制面板时出错（继续截图）: {e}")
        
        # 应用UI状态（同步当前浏览器中的复选框状态）
        print("[截图] 同步UI状态...")
        try:
            show_farthest = ui_state.get('showFarthestLine', True)
            show_distance_labels = ui_state.get('showDistanceLabels', True)
            
            apply_ui_state_script = f"""
            (function() {{
                try {{
                    // 设置最远直线开关
                    const farthestCheckbox = document.getElementById('toggleFarthest');
                    if (farthestCheckbox) {{
                        farthestCheckbox.checked = {str(show_farthest).lower()};
                        // 创建并触发change事件
                        const changeEvent = new Event('change', {{ bubbles: true }});
                        farthestCheckbox.dispatchEvent(changeEvent);
                        // 如果toggleFarthestLine函数存在，也调用它（双重保险）
                        if (typeof window.toggleFarthestLine === 'function') {{
                            window.toggleFarthestLine();
                        }}
                    }}
                    
                    // 设置距离标签开关
                    const distanceLabelsCheckbox = document.getElementById('toggleDistanceLabels');
                    if (distanceLabelsCheckbox) {{
                        distanceLabelsCheckbox.checked = {str(show_distance_labels).lower()};
                        // 创建并触发change事件
                        const changeEvent2 = new Event('change', {{ bubbles: true }});
                        distanceLabelsCheckbox.dispatchEvent(changeEvent2);
                        // 如果toggleDistanceLabels函数存在，也调用它（双重保险）
                        if (typeof window.toggleDistanceLabels === 'function') {{
                            window.toggleDistanceLabels();
                        }}
                    }}
                    
                    return true;
                }} catch(e) {{
                    console.error('应用UI状态时出错:', e);
                    return false;
                }}
            }})();
            """
            result = driver.execute_script(apply_ui_state_script)
            if result:
                print(f"[截图] ✓ UI状态已同步 (最远直线: {show_farthest}, 距离标签: {show_distance_labels})")
            else:
                print(f"[截图] ⚠️ UI状态同步可能失败，继续截图...")
            
            # 等待UI状态应用完成（给足够时间让地图重新渲染）
            time.sleep(1.0)
        except Exception as e:
            print(f"[截图] ⚠️ 应用UI状态时出错（继续截图）: {e}")
        
        # 执行JavaScript：滚动主页面（确保地图内容也显示完整）
        try:
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.2)
        except Exception as e:
            print(f"[截图] ⚠️ 滚动主页面时出错（忽略）: {e}")
        
        # 等待一小段时间，确保所有内容都渲染完成
        time.sleep(0.3)
        
        # 截图：只截取viewport（可视区域），不包括浏览器UI
        print("[截图] 正在截图（仅截取网页内容区域，不含浏览器UI）...")
        try:
            # 使用JavaScript获取viewport尺寸
            viewport_script = """
            return {
                width: window.innerWidth,
                height: window.innerHeight,
                scrollX: window.pageXOffset || window.scrollX,
                scrollY: window.pageYOffset || window.scrollY
            };
            """
            viewport = driver.execute_script(viewport_script)
            viewport_width = viewport['width']
            viewport_height = viewport['height']
            
            print(f"[截图] Viewport尺寸: {viewport_width}x{viewport_height}")
            
            # 截取整个浏览器窗口
            driver.save_screenshot(filepath)
            
            # 如果需要裁剪只保留viewport区域，可以使用PIL/Pillow
            # 但对于大多数情况，save_screenshot已经只截取viewport区域
            # 因为Selenium连接到远程调试端口时，截图只包含页面内容
            
            print(f"[截图] ✓ 截图已保存: {filepath}")
            
            # 检查文件是否真的创建了
            if not os.path.exists(filepath):
                raise Exception("截图文件未创建")
            
            file_size = os.path.getsize(filepath)
            if file_size == 0:
                raise Exception("截图文件为空")
            
            print(f"[截图] ✓ 文件大小: {file_size} 字节")
            return filepath
            
        except Exception as e:
            error_msg = f"截图失败: {str(e)}"
            print(f"[截图] ❌ {error_msg}")
            raise Exception(error_msg)
        
    except Exception as e:
        error_msg = f"截图过程出错: {str(e)}"
        print(f"[截图] ❌ {error_msg}")
        raise Exception(error_msg)
    finally:
        # 只关闭自己启动的浏览器实例，不关闭外部传入的
        if driver and should_close_driver:
            try:
                print("[截图] 正在关闭浏览器...")
                driver.quit()
                print("[截图] 浏览器已关闭")
            except Exception as e:
                print(f"[截图] ⚠️ 关闭浏览器时出错（忽略）: {e}")
        elif driver and not should_close_driver:
            print("[截图] 使用外部浏览器实例，不关闭浏览器")


def capture_screenshot_sync(url: str, save_dir: str = "网点图", wait_time: int = 2, ui_state: dict = None, driver_instance = None, group_name: str = "") -> str:
    """
    同步版本的截图函数（供Flask调用）
    注意：此函数已经是同步的，保留此函数以保持API兼容性
    
    Args:
        url: 要截图的网页URL
        save_dir: 保存目录，默认为"网点图"
        wait_time: 等待页面加载和滚动完成的时间（秒），默认2秒
        ui_state: UI状态字典，包含 showFarthestLine 和 showDistanceLabels 等状态
        driver_instance: 浏览器实例（可选）
        group_name: 网组名称，用于截图文件命名（可选）
    
    Returns:
        保存的文件路径
    """
    return capture_screenshot(url, save_dir, wait_time, ui_state, driver_instance, group_name)


if __name__ == "__main__":
    # 测试代码
    test_url = "http://127.0.0.1:5005"
    try:
        result = capture_screenshot_sync(test_url, wait_time=3)
        print(f"测试成功，截图保存到: {result}")
    except Exception as e:
        print(f"测试失败: {e}")
