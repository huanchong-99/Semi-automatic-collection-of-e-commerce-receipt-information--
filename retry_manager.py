#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重试管理器

负责管理验证码后的重试逻辑，包括配置管理、重试策略和状态跟踪
"""

import json
import os
import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

class RetryManager:
    """重试管理器类"""
    
    def __init__(self, config_file="retry_config.json"):
        self.config_file = config_file
        self.config = self._load_config()
        self.retry_attempts = {}
        self.retry_history = []
    
    def _load_config(self) -> Dict:
        """加载重试配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # 返回默认配置
                return self._get_default_config()
        except Exception as e:
            print(f"加载重试配置失败: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict:
        """获取默认配置"""
        return {
            "retry_settings": {
                "max_retry_attempts": 3,
                "retry_delay_seconds": 2.0,
                "use_coordinate_cache": True,
                "cache_expiry_hours": 24,
                "retry_strategies": [
                    "cached_coordinates",
                    "smart_element_search",
                    "fallback_xpath"
                ],
                "enable_retry_logging": True,
                "retry_timeout_seconds": 30
            },
            "coordinate_validation": {
                "enable_validation": True,
                "max_coordinate_age_minutes": 60,
                "min_success_count": 2,
                "screen_bounds_check": True
            },
            "fallback_options": {
                "enable_manual_intervention": True,
                "show_retry_dialog": False,
                "auto_skip_failed_elements": False
            }
        }
    
    def should_retry(self, element_name: str, order_index: int) -> bool:
        """判断是否应该重试"""
        key = f"{element_name}_{order_index}"
        attempts = self.retry_attempts.get(key, 0)
        max_attempts = self.config["retry_settings"]["max_retry_attempts"]
        
        return attempts < max_attempts
    
    def record_retry_attempt(self, element_name: str, order_index: int, success: bool) -> None:
        """记录重试尝试"""
        key = f"{element_name}_{order_index}"
        self.retry_attempts[key] = self.retry_attempts.get(key, 0) + 1
        
        # 记录重试历史
        self.retry_history.append({
            "timestamp": datetime.now().isoformat(),
            "element_name": element_name,
            "order_index": order_index,
            "attempt_number": self.retry_attempts[key],
            "success": success
        })
    
    def get_retry_delay(self) -> float:
        """获取重试延迟时间"""
        return self.config["retry_settings"]["retry_delay_seconds"]
    
    def get_retry_strategies(self) -> List[str]:
        """获取重试策略列表"""
        return self.config["retry_settings"]["retry_strategies"]
    
    def is_coordinate_cache_enabled(self) -> bool:
        """检查是否启用坐标缓存"""
        return self.config["retry_settings"]["use_coordinate_cache"]
    
    def is_coordinate_valid(self, coord_info: Dict) -> bool:
        """验证坐标是否有效"""
        if not self.config["coordinate_validation"]["enable_validation"]:
            return True
        
        try:
            # 检查成功次数
            min_success = self.config["coordinate_validation"]["min_success_count"]
            if coord_info.get("success_count", 0) < min_success:
                return False
            
            # 检查坐标年龄
            max_age_minutes = self.config["coordinate_validation"]["max_coordinate_age_minutes"]
            last_success = coord_info.get("last_success")
            if last_success:
                last_time = datetime.fromisoformat(last_success)
                age = datetime.now() - last_time
                if age.total_seconds() > max_age_minutes * 60:
                    return False
            
            # 检查屏幕边界（如果启用）
            if self.config["coordinate_validation"]["screen_bounds_check"]:
                screen_x = coord_info.get("screen_x", 0)
                screen_y = coord_info.get("screen_y", 0)
                if screen_x < 0 or screen_y < 0 or screen_x > 3000 or screen_y > 2000:
                    return False
            
            return True
            
        except Exception as e:
            print(f"坐标验证失败: {e}")
            return False
    
    def reset_retry_attempts(self, element_name: str = None, order_index: int = None) -> None:
        """重置重试计数"""
        if element_name and order_index is not None:
            key = f"{element_name}_{order_index}"
            if key in self.retry_attempts:
                del self.retry_attempts[key]
        else:
            self.retry_attempts.clear()
    
    def get_retry_statistics(self) -> Dict:
        """获取重试统计信息"""
        total_attempts = len(self.retry_history)
        successful_retries = sum(1 for h in self.retry_history if h["success"])
        
        return {
            "total_retry_attempts": total_attempts,
            "successful_retries": successful_retries,
            "success_rate": successful_retries / total_attempts if total_attempts > 0 else 0,
            "retry_history": self.retry_history[-10:],  # 最近10次重试记录
            "current_retry_counts": dict(self.retry_attempts)
        }
    
    def get_timeout_seconds(self) -> int:
        """获取重试超时时间"""
        return self.config["retry_settings"]["retry_timeout_seconds"]
    
    def is_manual_intervention_enabled(self) -> bool:
        """检查是否启用手动干预"""
        return self.config["fallback_options"]["enable_manual_intervention"]
    
    def should_show_retry_dialog(self) -> bool:
        """检查是否显示重试对话框"""
        return self.config["fallback_options"]["show_retry_dialog"]
    
    def should_auto_skip_failed_elements(self) -> bool:
        """检查是否自动跳过失败的元素"""
        return self.config["fallback_options"]["auto_skip_failed_elements"]
    
    def is_retry_logging_enabled(self) -> bool:
        """检查是否启用重试日志"""
        return self.config["retry_settings"]["enable_retry_logging"]
    
    def cleanup_old_history(self, max_history_size: int = 1000) -> None:
        """清理旧的重试历史记录"""
        if len(self.retry_history) > max_history_size:
            self.retry_history = self.retry_history[-max_history_size:]
    
    def export_statistics(self, filename: str = None) -> str:
        """导出重试统计信息"""
        if filename is None:
            filename = f"retry_statistics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        stats = self.get_retry_statistics()
        stats["export_time"] = datetime.now().isoformat()
        stats["config"] = self.config
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(stats, f, ensure_ascii=False, indent=2)
            return filename
        except Exception as e:
            print(f"导出统计信息失败: {e}")
            return None
    
    def update_config(self, new_config: Dict) -> bool:
        """更新配置"""
        try:
            self.config.update(new_config)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"更新配置失败: {e}")
            return False