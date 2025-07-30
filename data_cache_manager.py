#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据缓存管理器 - JSON缓存文件权限分离方案

实现统一的数据缓存，剪贴板监听器有写入权限，导出模块只有读取权限
确保数据关联的原子性和一致性
"""

import json
import os
import time
import threading
from datetime import datetime
from utils import *

class DataCacheManager:
    """数据缓存管理器 - 实现权限分离的缓存机制"""
    
    def __init__(self, cache_file_path="order_data_cache.json"):
        """初始化数据缓存管理器"""
        self.cache_file_path = cache_file_path
        self.write_lock = threading.Lock()  # 写入锁，防止并发冲突
        self.cache_data = {}
        self._load_cache()
    
    def _load_cache(self):
        """加载缓存文件"""
        try:
            if os.path.exists(self.cache_file_path):
                with open(self.cache_file_path, 'r', encoding='utf-8') as f:
                    self.cache_data = json.load(f)
                    if not isinstance(self.cache_data, dict):
                        self.cache_data = {}
            else:
                self.cache_data = {}
        except Exception as e:
            print(f"加载缓存文件失败: {str(e)}")
            self.cache_data = {}
    
    def _save_cache(self):
        """保存缓存到文件（原子性写入）"""
        try:
            # 使用临时文件确保原子性写入
            temp_file = self.cache_file_path + ".tmp"
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache_data, f, ensure_ascii=False, indent=2)
            
            # 原子性替换
            if os.path.exists(self.cache_file_path):
                os.remove(self.cache_file_path)
            os.rename(temp_file, self.cache_file_path)
            return True
        except Exception as e:
            print(f"保存缓存文件失败: {str(e)}")
            return False
    
    def write_order_data(self, order_id, order_data=None, shipping_info=None):
        """写入订单数据（剪贴板监听器专用 - 写入权限）"""
        with self.write_lock:
            clean_order_id = self._clean_order_id(order_id)
            if not clean_order_id or clean_order_id == "" or len(clean_order_id.strip()) == 0:
                print(f"[缓存拒绝] 无效订单ID: '{order_id}' -> '{clean_order_id}'")
                return False
            
            # 验证订单编号一致性
            if order_data and '订单编号' in order_data:
                order_num_field = order_data['订单编号']
                if order_num_field != clean_order_id:
                    print(f"[验证警告] 订单ID不匹配: key={clean_order_id}, field={order_num_field}")
                    # 修正订单编号字段
                    order_data['订单编号'] = clean_order_id
            
            # 检查重复收货信息（先创建副本避免迭代时修改字典）
            if shipping_info:
                cache_items = list(self.cache_data.items())  # 创建副本
                for existing_id, existing_data in cache_items:
                    if (existing_id != clean_order_id and 
                        existing_data.get('shipping_info') == shipping_info):
                        print(f"[重复警告] 发现重复收货信息: {clean_order_id} 与 {existing_id}")
                        # 标记为可能重复（稍后处理，避免在迭代中修改）
                        potential_duplicate = existing_id
                        break
                else:
                    potential_duplicate = None
            
            # 初始化订单记录
            if clean_order_id not in self.cache_data:
                self.cache_data[clean_order_id] = {
                    "order_id": clean_order_id,
                    "created_at": datetime.now().isoformat(),
                    "updated_at": datetime.now().isoformat(),
                    "status": "partial"
                }
            
            # 处理重复标记（在初始化后进行）
            if shipping_info and 'potential_duplicate' in locals() and potential_duplicate:
                self.cache_data[clean_order_id]['potential_duplicate'] = potential_duplicate
            
            # 更新订单基础数据
            if order_data:
                for key, value in order_data.items():
                    if key != "order_id":  # 避免覆盖标准化的order_id
                        self.cache_data[clean_order_id][key] = value
            
            # 更新收货信息
            if shipping_info:
                self.cache_data[clean_order_id]["shipping_info"] = shipping_info
                self.cache_data[clean_order_id]["shipping_info_updated_at"] = datetime.now().isoformat()
            
            # 更新状态
            self.cache_data[clean_order_id]["updated_at"] = datetime.now().isoformat()
            if shipping_info and order_data:
                self.cache_data[clean_order_id]["status"] = "completed"
            
            # 保存到文件
            success = self._save_cache()
            if success:
                status = self.cache_data[clean_order_id].get('status', 'unknown')
                print(f"[缓存写入] 订单ID: {clean_order_id}, 状态: {status}")
            return success
    
    def read_all_orders(self):
        """读取所有订单数据（导出模块专用 - 只读权限）"""
        # 重新加载最新数据
        self._load_cache()
        return dict(self.cache_data)  # 返回副本，防止意外修改
    
    def read_order_by_id(self, order_id):
        """根据订单ID读取单个订单数据（只读权限）"""
        clean_order_id = self._clean_order_id(order_id)
        if clean_order_id in self.cache_data:
            return dict(self.cache_data[clean_order_id])  # 返回副本
        return None
    
    def get_orders_with_shipping_info(self):
        """获取包含收货信息的订单（导出专用）"""
        self._load_cache()
        result = {}
        for order_id, data in self.cache_data.items():
            if "shipping_info" in data and data["shipping_info"]:
                result[order_id] = dict(data)  # 返回副本
        return result
    
    def get_cache_stats(self):
        """获取缓存统计信息"""
        self._load_cache()
        total_orders = len(self.cache_data)
        completed_orders = sum(1 for data in self.cache_data.values() if data.get("status") == "completed")
        partial_orders = total_orders - completed_orders
        
        return {
            "total_orders": total_orders,
            "completed_orders": completed_orders,
            "partial_orders": partial_orders,
            "cache_file": self.cache_file_path,
            "last_updated": max([data.get("updated_at", "") for data in self.cache_data.values()], default="")
        }
    
    def _clean_order_id(self, order_id):
        """清理订单ID格式"""
        if not order_id:
            return None
        
        # 移除常见前缀
        clean_id = str(order_id).strip()
        if clean_id.startswith("订单编号："):
            clean_id = clean_id.replace("订单编号：", "")
        elif clean_id.startswith("订单编号:"):
            clean_id = clean_id.replace("订单编号:", "")
        elif clean_id.startswith("："):  # 移除单独的中文冒号前缀
            clean_id = clean_id[1:]
        elif clean_id.startswith(":"):   # 移除单独的英文冒号前缀
            clean_id = clean_id[1:]
        
        return clean_id.strip() if clean_id.strip() else None
    
    def clear_cache(self):
        """清空缓存（谨慎使用）"""
        with self.write_lock:
            self.cache_data = {}
            return self._save_cache()
    
    def backup_cache(self, backup_suffix=None):
        """备份缓存文件"""
        if not backup_suffix:
            backup_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        backup_file = f"{self.cache_file_path}.backup_{backup_suffix}"
        try:
            if os.path.exists(self.cache_file_path):
                import shutil
                shutil.copy2(self.cache_file_path, backup_file)
                print(f"缓存已备份到: {backup_file}")
                return backup_file
        except Exception as e:
            print(f"备份缓存失败: {str(e)}")
        return None

# 全局缓存管理器实例
_cache_manager = None

def get_cache_manager():
    """获取全局缓存管理器实例"""
    global _cache_manager
    if _cache_manager is None:
        _cache_manager = DataCacheManager()
    return _cache_manager