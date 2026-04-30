#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAXA 参数化自动化脚本
功能：
1. 监控 CAXA 进程状态，必要时自动重启
2. 读取参数文件（JSON）并驱动 CAXA 绘图
3. 批量处理多个参数文件，自动保存图纸
4. 完整的错误捕获与日志，实现“智能纠错”
依赖：pywin32, psutil （安装命令见下方注释）
"""

import os
import json
import time
import sys
import logging
import traceback
from pathlib import Path
from typing import Optional, List, Dict, Any

import psutil
import win32com.client
import pythoncom

# ==================== 配置区 ====================
PROCESS_NAME = "CDRAFT_M.exe"
CAXA_PATH = r"C:\Program Files (x86)\CAXA\CAXA DRAFT MECHANICAL\2013\bin32\CDRAFT_M.exe"
PARAMS_DIR = "./params"          # 参数文件存放目录
OUTPUT_DIR = "./output"          # 生成的图纸保存目录
LOG_FILE = "caxa_auto.log"

# ==================== 日志配置 ====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# ==================== 进程管理 ====================
def is_caxa_running() -> bool:
    """检查CAXA是否正在运行"""
    for p in psutil.process_iter(['name']):
        if p.info['name'] and p.info['name'].lower() == PROCESS_NAME.lower():
            return True
    return False


def restart_caxa():
    """重启CAXA应用程序"""
    if is_caxa_running():
        for p in psutil.process_iter(['name', 'pid']):
            if p.info['name'] and p.info['name'].lower() == PROCESS_NAME.lower():
                psutil.Process(p.info['pid']).kill()
                logging.info(f"已终止卡死的CAXA进程 PID: {p.info['pid']}")
    subprocess.Popen([CAXA_PATH])
    time.sleep(8)  # 等待CAXA完全启动
    logging.info("CAXA 已重新启动")


# ==================== CAXA COM 操作核心 ====================
class CaxaAutomation:
    """CAXA自动化操作封装"""
    def __init__(self, simulate: bool = False):
        self.simulate = simulate
        self.app = None
        self.doc = None
        if not self.simulate:
            self._connect()

    def _connect(self):
        """连接CAXA COM服务"""
        try:
            # CAXA的ProgID通常是 "CAXA.Application" （版本不同可能有差异，常见为CAXA.Application.2013）
            self.app = win32com.client.Dispatch("CAXA.Application")
            logging.info("成功连接CAXA应用程序")
        except Exception:
            try:
                self.app = win32com.client.Dispatch("CAXA.Application.2013")
                logging.info("使用CAXA.Application.2013连接成功")
            except Exception as e:
                logging.error(f"无法连接CAXA: {e}")
                self.simulate = True
                logging.warning("将切换为模拟模式运行")

    def new_document(self):
        """新建图纸"""
        if self.simulate:
            logging.info("[模拟] 新建图纸")
            return
        self.doc = self.app.ActiveDocument
        if not self.doc:
            self.app.Documents.Add()
            self.doc = self.app.ActiveDocument
            logging.info("已新建图纸")

    def add_circle(self, x: float, y: float, radius: float):
        """绘制圆 (中心坐标单位mm)"""
        if self.simulate:
            logging.info(f"[模拟] 画圆: 中心({x},{y}) 半径{radius}")
            return
        try:
            # CAXA对象模型：AddCircle(CenterX, CenterY, Radius)
            self.doc.AddCircle(x, y, radius)
        except Exception as e:
            logging.error(f"画圆失败: {e}")

    def add_rect(self, x: float, y: float, width: float, height: float):
        """绘制矩形 (左下角点坐标,宽度,高度)"""
        if self.simulate:
            logging.info(f"[模拟] 画矩形: 左下({x},{y}) 宽{width} 高{height}")
            return
        try:
            # 使用直线绘制矩形
            self.doc.AddLine(x, y, x + width, y)
            self.doc.AddLine(x + width, y, x + width, y + height)
            self.doc.AddLine(x + width, y + height, x, y + height)
            self.doc.AddLine(x, y + height, x, y)
        except Exception as e:
            logging.error(f"画矩形失败: {e}")

    def add_text(self, x: float, y: float, text: str, height: float = 3.5):
        """添加单行文本"""
        if self.simulate:
            logging.info(f"[模拟] 添加文字: 位置({x},{y}) 内容'{text}'")
            return
        try:
            self.doc.AddText(x, y, text, height)
        except Exception as e:
            logging.error(f"添加文字失败: {e}")

    def save_as(self, filepath: str):
        """保存当前图纸"""
        if self.simulate:
            logging.info(f"[模拟] 保存图纸至: {filepath}")
            return
        self.doc.SaveAs(filepath)
        logging.info(f"已保存图纸: {filepath}")

    def close_document(self):
        """关闭当前文档"""
        if self.simulate or not self.doc:
            return
        self.doc.Close()


# ==================== 参数化模型处理 ====================
def load_params(file_path: str) -> Optional[Dict[str, Any]]:
    """加载单个参数文件（JSON格式），支持参数校验与纠错"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            params = json.load(f)
        # 必须包含的字段校验
        if 'type' not in params:
            logging.warning(f"参数文件缺少 'type' 字段: {file_path}，跳过")
            return None
        return params
    except json.JSONDecodeError as e:
        logging.error(f"参数文件格式错误 {file_path}: {e}")
        return None
    except Exception as e:
        logging.error(f"读取参数文件失败 {file_path}: {e}")
        return None


def execute_drawing(params: Dict[str, Any], caxa: CaxaAutomation):
    """根据参数执行绘图操作"""
    ptype = params['type']
    caxa.new_document()

    if ptype == 'circle':
        # 参数示例: {"type":"circle","x":50,"y":50,"radius":20}
        caxa.add_circle(params['x'], params['y'], params['radius'])
        caxa.add_text(params['x'], params['y'], f"R{params['radius']}")

    elif ptype == 'rectangle':
        # 参数示例: {"type":"rectangle","x":0,"y":0,"width":100,"height":50}
        caxa.add_rect(params['x'], params['y'], params['width'], params['height'])
        caxa.add_text(params['x'] + params['width']/2, params['y'] - 10,
                      f"{params['width']}x{params['height']}")

    elif ptype == 'bolt':
        # 简化螺栓：圆形头部 + 矩形螺杆
        head_d = params.get('head_d', 20)
        length = params.get('length', 60)
        # 画头部圆
        caxa.add_circle(0, 0, head_d/2)
        caxa.add_text(0, head_d/2 + 5, f"M{params.get('size', 8)}")
        # 画螺杆矩形
        caxa.add_rect(-head_d/4, -length, head_d/2, length)
    else:
        logging.warning(f"未知图形类型: {ptype}")


def batch_process(caxa: CaxaAutomation):
    """批量处理参数文件夹内的所有JSON文件"""
    params_dir = Path(PARAMS_DIR)
    if not params_dir.exists():
        logging.error(f"参数目录不存在: {PARAMS_DIR}")
        return

    files = list(params_dir.glob("*.json"))
    if not files:
        logging.warning("未找到任何参数文件")
        return

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    success_count = 0
    for f in files:
        logging.info(f"处理参数文件: {f.name}")
        params = load_params(f)
        if not params:
            continue

        try:
            execute_drawing(params, caxa)
            out_path = output_dir / f"{f.stem}.exb"
            caxa.save_as(str(out_path))
            success_count += 1
        except Exception as e:
            logging.error(f"绘图失败 [{f.name}]: {traceback.format_exc()}")
        finally:
            caxa.close_document()

    logging.info(f"批量处理完成: 成功 {success_count}/{len(files)}")


# ==================== 主菜单 ====================
def main():
    print("="*50)
    print("CAXA 参数化自动化脚本 V2.0")
    print("="*50)

    # 1. 检查CAXA运行状态，若未运行且未要求模拟则尝试启动
    simulate = False
    if not is_caxa_running():
        ans = input("CAXA 未运行，是否启动？(y/n, 直接回车为模拟模式): ").strip().lower()
        if ans == 'y':
            restart_caxa()
        else:
            simulate = True
            logging.info("将以模拟模式运行（不操作真实CAXA）")
    else:
        simulate = False

    # 2. 初始化自动化对象
    caxa = CaxaAutomation(simulate=simulate)

    # 3. 可选：执行进程守护（卡死检测）或者直接开始批量绘图
    print("\n功能选择:")
    print("1. 批量参数化绘图（读取 params 文件夹）")
    print("2. 进程守护模式（监控CAXA并自动重启）")
    print("3. 退出")
    choice = input("请选择: ").strip()

    if choice == '1':
        batch_process(caxa)
    elif choice == '2':
        # 简单守护循环，每30秒检查一次CAXA状态
        logging.info("进入CAXA进程守护模式，按 Ctrl+C 退出")
        try:
            while True:
                if not is_caxa_running():
                    logging.warning("CAXA 未运行，尝试重启...")
                    restart_caxa()
                    # 重启后可选择自动继续未完成的任务
                    batch_process(CaxaAutomation(simulate=False))
                time.sleep(30)
        except KeyboardInterrupt:
            logging.info("守护进程已手动停止")
    else:
        sys.exit()

if __name__ == "__main__":
    main()