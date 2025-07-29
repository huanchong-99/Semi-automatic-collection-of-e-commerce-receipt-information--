#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
坐标缓存系统

负责管理元素坐标的缓存和恢复，用于重试机制中的坐标回退功能
"""

import json
import os
import time
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple

class CoordinateCache:
    """坐标缓存管理器"""
    
    def __init__(self, cache_file="coordinate_cache.json"):
        self.cache_file = cache_file
        self.cache_data = self._load_cache()
        self.is_retry_mode = False  # 重试模式标志
        
    def _load_cache(self) -> Dict:
        """加载缓存数据"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return self._get_default_cache_structure()
        except Exception as e:
            print(f"加载坐标缓存失败: {e}")
            return self._get_default_cache_structure()
    
    def _get_default_cache_structure(self) -> Dict:
        """获取默认缓存结构"""
        return {
            "cache_version": "1.0",
            "last_updated": datetime.now().isoformat(),
            "coordinates": {}
        }
    
    def _save_cache(self) -> bool:
        """保存缓存数据"""
        try:
            self.cache_data["last_updated"] = datetime.now().isoformat()
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存坐标缓存失败: {e}")
            return False
    
    def set_retry_mode(self, is_retry: bool) -> None:
        """设置重试模式状态"""
        self.is_retry_mode = is_retry
        if not is_retry:
            print("[坐标缓存] 退出重试模式，停止使用缓存坐标")
    
    def should_use_cache(self) -> bool:
        """判断是否应该使用缓存坐标（仅在重试模式下）"""
        return self.is_retry_mode
    
    def save_coordinate(self, element_name: str, screen_x: int, screen_y: int, 
                      element_offset_x: int = 0, element_offset_y: int = 0) -> bool:
        """保存元素坐标（正常模式和重试模式都会保存）"""
        try:
            coordinate_data = {
                "screen_x": screen_x,
                "screen_y": screen_y,
                "element_offset_x": element_offset_x,
                "element_offset_y": element_offset_y,
                "success_count": self.cache_data["coordinates"].get(element_name, {}).get("success_count", 0) + 1,
                "last_success": datetime.now().isoformat(),
                "created_at": self.cache_data["coordinates"].get(element_name, {}).get("created_at", datetime.now().isoformat())
            }
            
            self.cache_data["coordinates"][element_name] = coordinate_data
            success = self._save_cache()
            
            if success:
                print(f"[坐标缓存] 已保存元素'{element_name}'的坐标: ({screen_x}, {screen_y})")
            
            return success
        except Exception as e:
            print(f"保存坐标失败: {e}")
            return False
    
    def get_cached_coordinate(self, element_name: str) -> Optional[Tuple[int, int]]:
        """获取缓存的坐标（仅在重试模式下且元素查找失败时使用）"""
        if not self.should_use_cache():
            return None
            
        try:
            if element_name in self.cache_data["coordinates"]:
                coord_data = self.cache_data["coordinates"][element_name]
                
                # 检查坐标有效性
                if self._is_coordinate_valid(coord_data):
                    screen_x = coord_data["screen_x"]
                    screen_y = coord_data["screen_y"]
                    print(f"[坐标缓存] 使用缓存坐标 '{element_name}': ({screen_x}, {screen_y})")
                    return (screen_x, screen_y)
                else:
                    print(f"[坐标缓存] 元素'{element_name}'的缓存坐标已过期或无效")
                    return None
            else:
                print(f"[坐标缓存] 未找到元素'{element_name}'的缓存坐标")
                return None
        except Exception as e:
            print(f"获取缓存坐标失败: {e}")
            return None
    
    def _is_coordinate_valid(self, coord_data: Dict) -> bool:
        """检查坐标是否有效"""
        try:
            # 检查必要字段
            required_fields = ["screen_x", "screen_y", "last_success"]
            for field in required_fields:
                if field not in coord_data:
                    return False
            
            # 检查时间有效性（24小时内）
            last_success = datetime.fromisoformat(coord_data["last_success"])
            if datetime.now() - last_success > timedelta(hours=24):
                return False
            
            # 检查成功次数（至少成功过1次）
            if coord_data.get("success_count", 0) < 1:
                return False
            
            # 检查坐标范围（简单的屏幕边界检查）
            x, y = coord_data["screen_x"], coord_data["screen_y"]
            if x < 0 or y < 0 or x > 3840 or y > 2160:  # 支持4K屏幕
                return False
            
            return True
        except Exception:
            return False
    
    def clear_expired_coordinates(self) -> int:
        """清理过期的坐标缓存"""
        try:
            expired_count = 0
            current_time = datetime.now()
            
            # 创建新的坐标字典，只保留有效的坐标
            new_coordinates = {}
            for element_name, coord_data in self.cache_data["coordinates"].items():
                if self._is_coordinate_valid(coord_data):
                    new_coordinates[element_name] = coord_data
                else:
                    expired_count += 1
            
            self.cache_data["coordinates"] = new_coordinates
            
            if expired_count > 0:
                self._save_cache()
                print(f"[坐标缓存] 已清理 {expired_count} 个过期坐标")
            
            return expired_count
        except Exception as e:
            print(f"清理过期坐标失败: {e}")
            return 0
    
    def get_cache_statistics(self) -> Dict:
        """获取缓存统计信息"""
        try:
            total_coordinates = len(self.cache_data["coordinates"])
            valid_coordinates = sum(1 for coord_data in self.cache_data["coordinates"].values() 
                                  if self._is_coordinate_valid(coord_data))
            
            return {
                "total_coordinates": total_coordinates,
                "valid_coordinates": valid_coordinates,
                "expired_coordinates": total_coordinates - valid_coordinates,
                "cache_file_size": os.path.getsize(self.cache_file) if os.path.exists(self.cache_file) else 0,
                "last_updated": self.cache_data.get("last_updated", "未知")
            }
        except Exception as e:
            print(f"获取缓存统计失败: {e}")
            return {}
    
    def reset_cache(self) -> bool:
        """重置缓存（清空所有坐标）"""
        try:
            self.cache_data = self._get_default_cache_structure()
            success = self._save_cache()
            if success:
                print("[坐标缓存] 缓存已重置")
            return success
        except Exception as e:
            print(f"重置缓存失败: {e}")
            return False