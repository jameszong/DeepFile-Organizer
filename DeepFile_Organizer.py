import gc
import io
import os
import re
import sys
import json
import time
import math
import threading
import asyncio
import hashlib
import traceback
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any, Set
from pathlib import Path
import pandas as pd
import numpy as np
import fitz  # PyMuPDF
import requests
from openai import OpenAI
from PIL import Image, ImageTk
import shutil
import sqlite3
from tqdm import tqdm
import pythoncom
import win32com.client
from docxtpl import DocxTemplate
import jinja2
# from rapidocr import RapidOCR  # 可选依赖，按需导入

# VS Code Dark+ Theme - 专业深色风格
COLORS = {
    'primary': '#007ACC',      # VS Code Blue
    'primary_dark': '#005FA3', # Darker Blue for hover
    'secondary': '#4EC9B0',    # VS Code Green (accent)
    'danger': '#F44747',       # VS Code Red
    'warning': '#CCA700',      # VS Code Yellow
    'bg_primary': '#1E1E1E',   # Main background
    'bg_secondary': '#252526', # Sidebar
    'bg_tertiary': '#2D2D30', # Panel/Popup
    'text_primary': '#CCCCCC', # Primary text
    'text_secondary': '#858585', # Muted text
    'text_tertiary': '#6A6A6A',  # Helper text
    'border': '#454545',       # Border color
    'input_bg': '#3C3C3C',     # Input/Button background
    'input_hover': '#4D4D4D',  # Input hover
    'selection': '#264F78',    # Selection highlight
    'folder': '#D7BA7D',       # Folder icon color
    'file': '#CCCCCC',         # File icon color
}

CONFIG_FILE = "config_settings.json"
HASH_CACHE_FILE = "ocr_hash_cache.json"

class HashCacheManager:
    """文件Hash缓存管理器 - 用于OCR结果缓存
    
    Phase 1: 使用JSON文件存储
    Phase 2: 迁移到SQLite数据库
    """
    def __init__(self, cache_file=HASH_CACHE_FILE):
        self.cache_file = cache_file
        self.cache = self._load_cache()
    
    def _load_cache(self) -> Dict:
        """加载缓存文件"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}
    
    def _save_cache(self):
        """保存缓存到文件"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[HashCache] 保存缓存失败: {e}")
    
    def compute_file_hash(self, file_path: str) -> str:
        """计算文件SHA256 Hash"""
        sha256 = hashlib.sha256()
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()
    
    def get_cached_ocr(self, file_hash: str, ocr_type: str = 'llm') -> Optional[Dict]:
        """获取缓存的OCR结果
        
        Args:
            file_hash: 文件Hash
            ocr_type: OCR类型 ('local' 或 'llm')
        
        Returns:
            缓存记录或None
        """
        if file_hash in self.cache:
            record = self.cache[file_hash]
            # 检查是否有指定类型的OCR结果
            if ocr_type in record.get('ocr_results', {}):
                return record
        return None
    
    def save_ocr_result(self, file_hash: str, file_path: str, ocr_text: str, 
                       ocr_type: str = 'llm', model_id: str = '', token_usage: int = 0):
        """保存OCR结果到缓存
        
        Args:
            file_hash: 文件Hash
            file_path: 文件路径
            ocr_text: OCR识别文本
            ocr_type: OCR类型 ('local' 或 'llm')
            model_id: LLM模型ID
            token_usage: Token使用量
        """
        if file_hash not in self.cache:
            self.cache[file_hash] = {
                'file_path': file_path,
                'first_seen': datetime.now().isoformat(),
                'ocr_results': {}
            }
        
        self.cache[file_hash]['ocr_results'][ocr_type] = {
            'text': ocr_text,
            'model_id': model_id,
            'token_usage': token_usage,
            'timestamp': datetime.now().isoformat()
        }
        
        self._save_cache()
    
    def clear_cache(self):
        """清空缓存"""
        self.cache = {}
        if os.path.exists(self.cache_file):
            try:
                os.remove(self.cache_file)
            except Exception:
                pass


class DatabaseManager:
    """SQLite数据库管理器 - 用于配置和OCR记录持久化
    
    Phase 2: 替代JSON文件存储
    """
    def __init__(self, db_file="deepfile.db"):
        self.db_file = db_file
        self._init_database()
    
    def _init_database(self):
        """初始化数据库表结构"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        # 配置表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS config (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # OCR记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS ocr_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_hash TEXT UNIQUE NOT NULL,
                file_path TEXT,
                ocr_text TEXT,
                ocr_source TEXT,
                model_id TEXT,
                token_usage INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 任务日志表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS task_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_type TEXT,
                start_time TIMESTAMP,
                end_time TIMESTAMP,
                files_processed INTEGER DEFAULT 0,
                files_matched INTEGER DEFAULT 0,
                total_tokens INTEGER DEFAULT 0,
                status TEXT
            )
        ''')
        
        # ========== Tab5 向量归档专用表 ==========
        
        # 文件元数据表（存储向量和OCR结果）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS file_metadata (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_hash TEXT UNIQUE NOT NULL,
                file_path TEXT,
                file_name TEXT,
                file_type TEXT,
                pdf_metadata TEXT,
                ocr_text TEXT,
                ocr_source TEXT,
                ocr_model_id TEXT,
                ocr_token_usage INTEGER DEFAULT 0,
                content_vector BLOB,
                name_vector BLOB,
                meta_vector BLOB,
                vector_token_usage INTEGER DEFAULT 0,
                is_classified BOOLEAN DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 分类结果表（存储匹配结果和冲突标记）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS classification_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER,
                file_hash TEXT,
                row_index INTEGER,
                match_score REAL,
                matched_columns TEXT,
                weight_details TEXT,
                is_conflicted BOOLEAN DEFAULT 0,
                final_decision TEXT,
                target_folder TEXT,
                target_filename TEXT,
                status TEXT DEFAULT 'pending',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (file_hash) REFERENCES file_metadata(file_hash)
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def get_config(self, key: str, default=None):
        """获取配置值"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM config WHERE key = ?", (key,))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else default
    
    def set_config(self, key: str, value: str):
        """设置配置值"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO config (key, value, updated_at)
            VALUES (?, ?, CURRENT_TIMESTAMP)
        ''', (key, value))
        conn.commit()
        conn.close()
    
    def save_ocr_record(self, file_hash: str, file_path: str, ocr_text: str,
                       ocr_source: str = 'llm', model_id: str = '', token_usage: int = 0):
        """保存OCR记录"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO ocr_records 
            (file_hash, file_path, ocr_text, ocr_source, model_id, token_usage, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (file_hash, file_path, ocr_text, ocr_source, model_id, token_usage))
        conn.commit()
        conn.close()
    
    def get_ocr_record(self, file_hash: str, ocr_source: str = 'llm'):
        """获取OCR记录"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT file_path, ocr_text, model_id, token_usage, created_at
            FROM ocr_records 
            WHERE file_hash = ? AND ocr_source = ?
        ''', (file_hash, ocr_source))
        result = cursor.fetchone()
        conn.close()
        if result:
            return {
                'file_path': result[0],
                'ocr_text': result[1],
                'model_id': result[2],
                'token_usage': result[3],
                'timestamp': result[4]
            }
        return None
    
    def log_task_start(self, task_type: str) -> int:
        """记录任务开始，返回任务ID"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO task_logs (task_type, start_time, status)
            VALUES (?, CURRENT_TIMESTAMP, 'running')
        ''', (task_type,))
        task_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return task_id
    
    def log_task_end(self, task_id: int, status: str, files_processed: int = 0,
                    files_matched: int = 0, total_tokens: int = 0):
        """记录任务结束"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE task_logs 
            SET end_time = CURRENT_TIMESTAMP, status = ?,
                files_processed = ?, files_matched = ?, total_tokens = ?
            WHERE id = ?
        ''', (status, files_processed, files_matched, total_tokens, task_id))
        conn.commit()
        conn.close()
    
    def migrate_from_json(self, json_file: str = "config_settings.json"):
        """从JSON文件迁移数据到SQLite"""
        if os.path.exists(json_file):
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                for key, value in config_data.items():
                    self.set_config(key, str(value))
                return True
            except Exception as e:
                print(f"[Database] 迁移JSON数据失败: {e}")
                return False
        return False

    # ========== Tab5 向量归档数据库操作 ==========
    
    def save_file_metadata(self, file_hash: str, file_path: str, file_name: str,
                          file_type: str, pdf_metadata: str, ocr_text: str,
                          ocr_source: str, ocr_model_id: str, ocr_token_usage: int,
                          content_vector: bytes, name_vector: bytes, meta_vector: bytes,
                          vector_token_usage: int):
        """保存文件元数据和向量到数据库"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO file_metadata 
            (file_hash, file_path, file_name, file_type, pdf_metadata, ocr_text,
             ocr_source, ocr_model_id, ocr_token_usage, content_vector, name_vector,
             meta_vector, vector_token_usage, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (file_hash, file_path, file_name, file_type, pdf_metadata, ocr_text,
             ocr_source, ocr_model_id, ocr_token_usage, content_vector, name_vector,
             meta_vector, vector_token_usage))
        conn.commit()
        conn.close()
    
    def get_file_metadata(self, file_hash: str):
        """获取文件元数据"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT file_path, file_name, file_type, pdf_metadata, ocr_text,
                   ocr_source, ocr_model_id, ocr_token_usage, content_vector,
                   name_vector, meta_vector, vector_token_usage, is_classified
            FROM file_metadata WHERE file_hash = ?
        ''', (file_hash,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return {
                'file_path': result[0],
                'file_name': result[1],
                'file_type': result[2],
                'pdf_metadata': result[3],
                'ocr_text': result[4],
                'ocr_source': result[5],
                'ocr_model_id': result[6],
                'ocr_token_usage': result[7],
                'content_vector': result[8],
                'name_vector': result[9],
                'meta_vector': result[10],
                'vector_token_usage': result[11],
                'is_classified': result[12]
            }
        return None
    
    def get_unclassified_files(self):
        """获取所有未分类的文件"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT file_hash, file_path, file_name, file_type, pdf_metadata, ocr_text,
                   content_vector, name_vector, meta_vector
            FROM file_metadata WHERE is_classified = 0
        ''')
        results = cursor.fetchall()
        conn.close()
        return [{
            'file_hash': r[0],
            'file_path': r[1],
            'file_name': r[2],
            'file_type': r[3],
            'pdf_metadata': r[4],
            'ocr_text': r[5],
            'content_vector': r[6],
            'name_vector': r[7],
            'meta_vector': r[8]
        } for r in results]
    
    def save_classification_result(self, task_id: int, file_hash: str, row_index: int,
                                   match_score: float, matched_columns: str, weight_details: str,
                                   is_conflicted: bool, target_folder: str, target_filename: str):
        """保存分类结果"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO classification_results 
            (task_id, file_hash, row_index, match_score, matched_columns, weight_details,
             is_conflicted, target_folder, target_filename, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending')
        ''', (task_id, file_hash, row_index, match_score, matched_columns, weight_details,
             is_conflicted, target_folder, target_filename))
        result_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return result_id
    
    def update_file_classified_status(self, file_hash: str, is_classified: bool = True):
        """更新文件分类状态"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE file_metadata SET is_classified = ?, updated_at = CURRENT_TIMESTAMP
            WHERE file_hash = ?
        ''', (is_classified, file_hash))
        conn.commit()
        conn.close()
    
    def update_classification_decision(self, result_id: int, final_decision: str, status: str = 'approved'):
        """更新分类决策"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE classification_results 
            SET final_decision = ?, status = ?
            WHERE id = ?
        ''', (final_decision, status, result_id))
        conn.commit()
        conn.close()
    
    def get_task_token_usage(self, task_id: int):
        """获取任务的总token使用情况"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT SUM(fm.ocr_token_usage), SUM(fm.vector_token_usage)
            FROM file_metadata fm
            JOIN classification_results cr ON fm.file_hash = cr.file_hash
            WHERE cr.task_id = ?
        ''', (task_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return {
                'ocr_tokens': result[0] or 0,
                'vector_tokens': result[1] or 0,
                'total_tokens': (result[0] or 0) + (result[1] or 0)
            }
        return {'ocr_tokens': 0, 'vector_tokens': 0, 'total_tokens': 0}
    
    def save_log(self, timestamp: str, message: str):
        """保存日志到数据库"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        # 创建日志表（如果不存在）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                message TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 插入日志记录
        cursor.execute('''
            INSERT INTO logs (timestamp, message)
            VALUES (?, ?)
        ''', (timestamp, message))
        
        conn.commit()
        conn.close()
    
    def get_recent_logs(self, limit: int = 100):
        """获取最近的日志记录"""
        import sqlite3
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT timestamp, message FROM logs 
            ORDER BY created_at DESC 
            LIMIT ?
        ''', (limit,))
        results = cursor.fetchall()
        conn.close()
        return results


class RateLimiter:
    """异步限流控制器 - 支持RPM和TPM限流
    
    配置参数:
    - rpm: 每分钟最大请求数 (requests per minute)
    - tpm: 每分钟最大token数 (tokens per minute)
    """
    def __init__(self, rpm: int = 5000, tpm: int = 1200000):
        self.rpm = rpm
        self.tpm = tpm
        self.request_times = []  # 请求时间戳列表
        self.token_usage = []    # token使用量列表 [(timestamp, tokens), ...]
        self.lock = asyncio.Lock()
        
        # 计算窗口大小（秒）
        self.window_size = 60  # 1分钟窗口
        
    async def acquire(self, estimated_tokens: int = 0):
        """获取执行许可，如果超过限流则等待
        
        Args:
            estimated_tokens: 预估消耗的token数（用于TPM限流）
        """
        async with self.lock:
            now = time.time()
            
            # 清理过期记录（超过1分钟的）
            cutoff = now - self.window_size
            self.request_times = [t for t in self.request_times if t > cutoff]
            self.token_usage = [(t, tokens) for t, tokens in self.token_usage if t > cutoff]
            
            # 计算当前窗口内的统计
            current_requests = len(self.request_times)
            current_tokens = sum(tokens for _, tokens in self.token_usage)
            
            # 检查RPM限流（预留10%安全余量）
            rpm_limit = self.rpm * 0.9
            if current_requests >= rpm_limit:
                # 计算需要等待的时间
                oldest_request = min(self.request_times)
                wait_time = self.window_size - (now - oldest_request) + 0.1
                if wait_time > 0:
                    print(f"[RateLimiter] RPM限流，等待 {wait_time:.2f} 秒...")
                    await asyncio.sleep(wait_time)
            
            # 检查TPM限流（预留10%安全余量）
            tpm_limit = self.tpm * 0.9
            if current_tokens + estimated_tokens > tpm_limit:
                # 计算需要等待的时间
                if self.token_usage:
                    oldest_token_time = min(t for t, _ in self.token_usage)
                    wait_time = self.window_size - (now - oldest_token_time) + 0.1
                    if wait_time > 0:
                        print(f"[RateLimiter] TPM限流，等待 {wait_time:.2f} 秒...")
                        await asyncio.sleep(wait_time)
            
            # 记录本次请求
            self.request_times.append(time.time())
            if estimated_tokens > 0:
                self.token_usage.append((time.time(), estimated_tokens))
    
    def release(self, actual_tokens: int = 0):
        """释放并记录实际使用的token数"""
        # 更新最后一次请求的token使用量
        if self.token_usage and actual_tokens > 0:
            last_time, _ = self.token_usage[-1]
            self.token_usage[-1] = (last_time, actual_tokens)


class VectorManager:
    """向量管理器 - 处理Qwen Embedding API调用和向量计算
    
    功能:
    - 调用Qwen Embedding API获取文本向量
    - 计算余弦相似度
    - 序列化/反序列化向量
    - Token使用量追踪
    """
    
    def __init__(self, api_key: str, model: str = "text-embedding-v3"):
        self.api_key = api_key
        self.model = model
        self.embedding_dim = 1024
        self.api_url = "https://dashscope.aliyuncs.com/compatible-mode/v1/embeddings"
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        })
    
    def get_embedding(self, texts: list) -> tuple:
        """获取文本的嵌入向量
        
        Args:
            texts: 文本列表（最多10个）
        
        Returns:
            (vectors, token_usage) - 向量和token使用量
        """
        if not texts:
            return [], 0
        
        # 限制批次大小
        if len(texts) > 10:
            results = []
            total_tokens = 0
            for i in range(0, len(texts), 10):
                batch = texts[i:i+10]
                vectors, tokens = self._batch_embedding(batch)
                results.extend(vectors)
                total_tokens += tokens
            return results, total_tokens
        
        return self._batch_embedding(texts)
    
    def _batch_embedding(self, texts: list) -> tuple:
        """获取向量嵌入内部方法"""
        try:
            payload = {
                "model": self.model,
                "input": texts,
                "encoding_format": "float"
            }
            response = self.session.post(self.api_url, json=payload, timeout=30)
            response.raise_for_status()
            data = response.json()

            vectors = []
            for item in data.get('data', []):
                embedding = item.get('embedding')
                if embedding:
                    vectors.append(np.array(embedding, dtype=np.float32))

            usage = data.get('usage', {})
            token_usage = usage.get('total_tokens', 0)

            if vectors:
                self.embedding_dim = len(vectors[0])
                return vectors, token_usage

            raise ValueError(f"DashScope返回中未包含embedding数据: {data}")
            return vectors, token_usage
        except Exception as e:
            print(f"API调用失败: {e}")
            return [np.zeros(self.embedding_dim, dtype=np.float32) for _ in texts], 0

    @staticmethod
    def cosine_similarity(vec1: np.ndarray, vec2: np.ndarray) -> float:
        """计算两个向量的余弦相似度"""
        if vec1 is None or vec2 is None:
            return 0.0
        
        norm1 = np.linalg.norm(vec1)
        norm2 = np.linalg.norm(vec2)
        
        if norm1 == 0 or norm2 == 0:
            return 0.0
        
        return float(np.dot(vec1, vec2) / (norm1 * norm2))
    
    @staticmethod
    def vector_to_bytes(vector: np.ndarray) -> bytes:
        """将numpy向量转换为bytes（用于数据库存储）"""
        if vector is None:
            return None
        return vector.astype(np.float32).tobytes()
    
    @staticmethod
    def bytes_to_vector(data: bytes) -> np.ndarray:
        """将bytes转换回numpy向量"""
        if data is None:
            return None
        return np.frombuffer(data, dtype=np.float32)
    
    def calculate_match_score(self, file_data: dict, key_values: dict, 
                             base_weights: dict, source_coefficients: dict) -> tuple:
        """计算文件与Excel行的匹配分数
        
        Args:
            file_data: 包含向量的文件数据
                - content_vector: 内容向量
                - name_vector: 文件名向量
                - meta_vector: PDF元数据向量
            key_values: 关键列值字典
                {列名: 值}
            base_weights: 关键列基础权重
                {列名: 权重值}
            source_coefficients: 来源系数
                {'content': 1.0, 'name': 1.5, 'meta': 0.7}
        
        Returns:
            (total_score, matched_columns, weight_details)
        """
        total_score = 0
        matched_columns = []
        weight_details = []
        
        # 反序列化向量
        content_vec = file_data.get('content_vector')
        name_vec = file_data.get('name_vector')
        meta_vec = file_data.get('meta_vector')
        
        # 如果数据库返回的是bytes，需要转换
        if isinstance(content_vec, bytes):
            content_vec = self.bytes_to_vector(content_vec)
        if isinstance(name_vec, bytes):
            name_vec = self.bytes_to_vector(name_vec)
        if isinstance(meta_vec, bytes):
            meta_vec = self.bytes_to_vector(meta_vec)
        
        # 获取关键列的向量
        key_vectors = []
        for key_name, key_value in key_values.items():
            if not key_value:
                continue
            vec, _ = self.get_embedding([key_value])
            if vec:
                key_vectors.append((key_name, vec[0], key_value))
        
        # 计算各来源的相似度
        for key_name, key_vec, key_value in key_vectors:
            base_weight = base_weights.get(key_name, 50)
            
            # 内容匹配
            if content_vec is not None:
                sim = self.cosine_similarity(key_vec, content_vec)
                if sim >= 0.75:  # 相似度阈值
                    score = base_weight * source_coefficients['content'] * sim
                    total_score += score
                    matched_columns.append(key_name)
                    weight_details.append({
                        'column': key_name,
                        'value': key_value,
                        'source': 'content',
                        'similarity': round(sim, 4),
                        'score': round(score, 2)
                    })
                    continue
            
            # 文件名匹配
            if name_vec is not None:
                sim = self.cosine_similarity(key_vec, name_vec)
                if sim >= 0.75:
                    score = base_weight * source_coefficients['name'] * sim
                    total_score += score
                    matched_columns.append(key_name)
                    weight_details.append({
                        'column': key_name,
                        'value': key_value,
                        'source': 'name',
                        'similarity': round(sim, 4),
                        'score': round(score, 2)
                    })
                    continue
            
            # PDF元数据匹配
            if meta_vec is not None:
                sim = self.cosine_similarity(key_vec, meta_vec)
                if sim >= 0.75:
                    score = base_weight * source_coefficients['meta'] * sim
                    total_score += score
                    matched_columns.append(key_name)
                    weight_details.append({
                        'column': key_name,
                        'value': key_value,
                        'source': 'meta',
                        'similarity': round(sim, 4),
                        'score': round(score, 2)
                    })
        
        return total_score, matched_columns, weight_details


class FileToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepFile Organizer - Enterprise Edition")
        self.root.geometry("1000x800")
        self.root.configure(bg=COLORS['bg_primary'])

        # 初始化
        self.ocr_engine = None
        self.tab4_ocr_engine = None
        self.tab5_ocr_engine = None
        self.config = self.load_config()
        self.db = DatabaseManager()  # 初始化数据库管理器

        # 主布局 - 使用PanedWindow确保日志区域固定高度
        self.main_paned = tk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_paned.pack(fill="both", expand=True, padx=12, pady=12)

        # 上部：Notebook区域
        self.notebook_frame = tk.Frame(self.main_paned, bg=COLORS['bg_primary'])
        self.main_paned.add(self.notebook_frame, height=550)  # 固定上部高度

        # Notebook样式配置 - VS Code风格标签
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TNotebook', 
                           background=COLORS['bg_secondary'], 
                           tabmargins=[2, 5, 2, 0])
        self.style.configure('TNotebook.Tab', 
                           background=COLORS['bg_secondary'], 
                           foreground=COLORS['text_secondary'],
                           padding=[12, 6],
                           font=('微软雅黑', 10))
        self.style.map('TNotebook.Tab', 
                      background=[('selected', COLORS['bg_primary'])],
                      foreground=[('selected', COLORS['text_primary'])],
                      expand=[('selected', [2, 2, 2, 0])])
        
        # 移除选中Tab的内陷效果 - VS Code风格平铺
        self.style.layout('TNotebook.Tab', [
            ('Notebook.tab', {
                'sticky': 'nswe', 
                'children': [
                    ('Notebook.padding', {
                        'side': 'top', 
                        'sticky': 'nswe',
                        'children': [
                            ('Notebook.label', {'side': 'left', 'sticky': ''})
                        ]
                    })
                ]
            })
        ])
        
        # 卡片样式
        self.style.configure('Card.TLabelframe', 
                           background=COLORS['bg_primary'],
                           borderwidth=1,
                           relief='solid')
        self.style.configure('Card.TLabelframe.Label',
                           background=COLORS['bg_primary'],
                           foreground=COLORS['text_primary'],
                           font=('微软雅黑', 10, 'bold'))

        self.notebook = ttk.Notebook(self.notebook_frame)
        self.notebook.pack(fill="both", expand=True)

        self.tab1 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)
        self.tab2 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)
        self.tab3 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)
        self.tab4 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)
        self.tab5 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)
        self.tab6 = tk.Frame(self.notebook, bg=COLORS['bg_primary'], padx=16, pady=16)

        self.notebook.add(self.tab1, text="📁 文件批量提取")
        self.notebook.add(self.tab2, text="📄 PDF智能重命名")
        self.notebook.add(self.tab3, text="🗂️ 高级文件整理")
        self.notebook.add(self.tab4, text="📂 智能文件归档")
        self.notebook.add(self.tab5, text="🧠 向量智能归档")
        self.notebook.add(self.tab6, text="📝 文件批量填充")

        # 下部：日志区域 - 固定高度
        self.log_frame = tk.LabelFrame(self.main_paned, 
                                     text="系统日志", 
                                     bg=COLORS['bg_secondary'],
                                     fg=COLORS['text_primary'],
                                     font=('微软雅黑', 10, 'bold'),
                                     padx=8, pady=8)
        self.main_paned.add(self.log_frame, height=180)  # 固定日志高度
        
        # 日志控制栏
        log_control_frame = tk.Frame(self.log_frame, bg=COLORS['bg_secondary'])
        log_control_frame.pack(fill="x", side="top", pady=(0, 5))
        
        tk.Label(log_control_frame, 
                text="操作记录", 
                bg=COLORS['bg_secondary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        
        tk.Button(log_control_frame, 
                 text="🗑️ 清空日志", 
                 command=self.clear_log,
                 bg=COLORS['bg_secondary'],
                 fg=COLORS['text_primary'],
                 relief='flat',
                 font=('微软雅黑', 9)).pack(side="right", padx=5)
        
        tk.Button(log_control_frame, 
                 text="💾 导出日志", 
                 command=self.export_log,
                 bg=COLORS['bg_secondary'],
                 fg=COLORS['text_primary'],
                 relief='flat',
                 font=('微软雅黑', 9)).pack(side="right", padx=5)

        # 日志文本区域
        self.log_text = scrolledtext.ScrolledText(
            self.log_frame, 
            height=8,  # 固定行数
            state='disabled', 
            font=("Consolas", 9),
            bg=COLORS['bg_primary'],
            fg=COLORS['text_primary'],
            relief='flat',
            padx=8,
            pady=8
        )
        self.log_text.pack(fill="both", expand=True)

        self.setup_tab1()
        self.setup_tab2()
        self.setup_tab3()
        self.setup_tab4()
        self.setup_tab5()
        self.setup_tab6()
        
        self.log("DeepFile Organizer 已启动 - Enterprise Edition")
        self.log("基于 Minimalism + Enterprise 设计规范重构")
        
        # 加载最近的日志
        self._load_recent_logs()

    # ================= 工具函数 =================
    def _load_recent_logs(self):
        """加载最近的日志记录"""
        try:
            if hasattr(self, 'db') and self.db:
                recent_logs = self.db.get_recent_logs(50)
                for timestamp, message in reversed(recent_logs):
                    self.log_text.config(state='normal')
                    self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state='disabled')
        except Exception as e:
            pass  # 如果加载日志失败，不影响程序启动
    
    def create_card(self, parent, title):
        """创建Enterprise风格卡片容器"""
        card = tk.LabelFrame(parent, 
                            text=title,
                            bg=COLORS['bg_primary'],
                            fg=COLORS['text_primary'],
                            font=('微软雅黑', 10, 'bold'),
                            padx=12, pady=8,
                            relief='solid',
                            borderwidth=1)
        return card

    def create_primary_button(self, parent, text, command):
        """创建主操作按钮 - Enterprise风格"""
        return tk.Button(parent,
                        text=text,
                        command=command,
                        bg=COLORS['primary'],
                        fg='white',
                        font=('微软雅黑', 10, 'bold'),
                        relief='flat',
                        padx=16,
                        pady=6,
                        cursor='hand2')

    def create_secondary_button(self, parent, text, command):
        """创建次要操作按钮"""
        return tk.Button(parent,
                        text=text,
                        command=command,
                        bg=COLORS['bg_tertiary'],
                        fg=COLORS['text_primary'],
                        font=('微软雅黑', 9),
                        relief='flat',
                        padx=12,
                        pady=4,
                        cursor='hand2')

    def create_centered_dialog(self, title, width, height):
        """创建居中模态对话框，带遮罩效果防止误触底层"""
        # 创建遮罩层 - 使用深色Frame覆盖整个主窗口
        overlay = tk.Frame(self.root, bg='#000000')
        overlay.place(x=0, y=0, relwidth=1, relheight=1)
        # 使用深色但不透明，模拟遮罩效果
        overlay.configure(bg='#1E1E1E')
        
        # 阻止所有与底层窗口的交互
        overlay.bind("<Button-1>", lambda e: "break")
        overlay.bind("<Key>", lambda e: "break")
        
        # 创建对话框窗口
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_tertiary'])
        
        # 移除窗口装饰（可选，看起来更简洁）
        dialog.overrideredirect(False)
        
        # 计算居中位置
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        dialog_x = root_x + (root_width - width) // 2
        dialog_y = root_y + (root_height - height) // 2
        dialog.geometry(f"{width}x{height}+{dialog_x}+{dialog_y}")
        
        # 对话框关闭时移除遮罩
        def on_dialog_close():
            overlay.destroy()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        
        # 点击遮罩关闭对话框
        def on_overlay_click(event):
            on_dialog_close()
        
        overlay.bind("<Button-1>", on_overlay_click)
        
        # 设置对话框为模态，确保它在最前面
        dialog.focus_set()
        dialog.lift()
        
        # 返回对话框和关闭函数，方便调用者使用
        return dialog, on_dialog_close, overlay

    # --- 配置管理 ---
    def load_config(self):
        default = {"api_key": "", "model_id": "doubao-seed-1-6-vision-250815", "ocr_mode": "local"}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return {**default, **json.load(f)}
            except: pass
        return default

    def save_config(self):
        self.config["api_key"] = self.api_key_var.get()
        self.config["model_id"] = self.model_id_var.get()
        self.config["ocr_mode"] = self.ocr_mode_var.get()
        
        # Tab5 配置
        self.config["tab5_vector_api_key"] = self.tab5_vector_api_key_var.get()
        self.config["tab5_vector_model"] = self.tab5_vector_model_var.get()
        self.config["tab5_ocr_api_key"] = self.tab5_ocr_api_key_var.get()
        self.config["tab5_ocr_model"] = self.tab5_ocr_model_var.get()
        self.config["tab5_use_tab4_ocr"] = self.tab5_use_tab4_ocr_var.get()
        self.config["tab5_folder_name_col1"] = self.tab5_folder_name_col1_var.get()
        self.config["tab5_folder_name_col2"] = self.tab5_folder_name_col2_var.get()
        self.config["tab5_folder_name_col3"] = self.tab5_folder_name_col3_var.get()
        self.config["tab5_file_name_col1"] = self.tab5_file_name_col1_var.get()
        self.config["tab5_file_name_col2"] = self.tab5_file_name_col2_var.get()
        self.config["tab5_file_name_col3"] = self.tab5_file_name_col3_var.get()
        
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def log(self, message):
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()
        
        # 持久化日志到数据库
        try:
            if hasattr(self, 'db') and self.db:
                self.db.save_log(timestamp, message)
        except Exception as e:
            # 如果保存日志失败，不影响主要功能
            pass

    def clear_log(self):
        """清空日志内容"""
        # 清空界面显示
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # 清空数据库中的日志
        try:
            if hasattr(self, 'db') and self.db:
                import sqlite3
                conn = sqlite3.connect(self.db.db_file)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM logs")
                conn.commit()
                conn.close()
                self.log("日志已清空（数据库已删除）")
        except Exception as e:
            self.log(f"清空数据库日志失败: {e}")

    def export_log(self):
        """导出日志到文件"""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            if not log_content.strip():
                messagebox.showinfo("提示", "日志内容为空，无需导出")
                return
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            default_filename = f"log_export_{timestamp}.txt"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                initialfile=default_filename,
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"DeepFile Organizer - 日志导出\n")
                    f.write(f"导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(log_content)
                
                self.log(f"日志已导出到: {file_path}")
                messagebox.showinfo("导出成功", f"日志已保存到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出日志时出错:\n{e}")

    # ================= Tab 1: 文件批量提取 (Enterprise风格) =================
    def setup_tab1(self):
        self.src_path_var = tk.StringVar()
        self.dest_path_var = tk.StringVar()
        self.rename_by_folder_var = tk.BooleanVar(value=False)

        frame = self.tab1
        frame.configure(bg=COLORS['bg_primary'])

        # 标题
        title_frame = tk.Frame(frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="文件批量提取",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="从源文件夹按条件筛选并复制文件到目标位置",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))

        # 1. 源文件夹卡片
        src_card = self.create_card(frame, "源文件夹")
        src_card.pack(fill="x", pady=8)
        
        src_row = tk.Frame(src_card, bg=COLORS['bg_primary'])
        src_row.pack(fill="x")
        tk.Entry(src_row, 
                textvariable=self.src_path_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(src_row, "浏览...", lambda: self.select_dir(self.src_path_var)).pack(side="right")

        # 2. 筛选条件卡片
        filter_card = self.create_card(frame, "筛选条件")
        filter_card.pack(fill="x", pady=8)

        filter_grid = tk.Frame(filter_card, bg=COLORS['bg_primary'])
        filter_grid.pack(fill="x")

        # 文件名模式
        tk.Label(filter_grid, 
                text="文件名包含:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="w", pady=6)
        self.name_pattern_entry = tk.Entry(filter_grid, width=45, 
                                          bg=COLORS['input_bg'],
                                          fg=COLORS['text_primary'],
                                          insertbackground=COLORS['text_primary'],
                                          relief='solid', highlightthickness=1, highlightcolor=COLORS['border'])
        self.name_pattern_entry.insert(0, "*仲裁申请书*") 
        self.name_pattern_entry.grid(row=0, column=1, padx=(12, 0), sticky="w")

        # 后缀名
        tk.Label(filter_grid, 
                text="文件后缀:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="w", pady=6)
        self.ext_entry = tk.Entry(filter_grid, width=45, 
                                bg=COLORS['input_bg'],
                                fg=COLORS['text_primary'],
                                insertbackground=COLORS['text_primary'],
                                relief='solid', highlightthickness=1, highlightcolor=COLORS['border'])
        self.ext_entry.insert(0, "pdf, doc, docx")
        self.ext_entry.grid(row=1, column=1, padx=(12, 0), sticky="w")

        # 冲突处理选项
        conflict_frame = tk.Frame(filter_card, bg=COLORS['bg_primary'])
        conflict_frame.pack(fill="x", pady=(12, 0))
        tk.Checkbutton(conflict_frame, 
                      text="重命名冲突文件：[来源文件夹名]_[原文件名]",
                      variable=self.rename_by_folder_var,
                      bg=COLORS['bg_primary'],
                      fg=COLORS['text_primary'],
                      selectcolor=COLORS['bg_primary'],
                      font=('微软雅黑', 9)).pack(anchor="w")

        # 3. 目标文件夹卡片
        dest_card = self.create_card(frame, "目标文件夹")
        dest_card.pack(fill="x", pady=8)
        
        dest_row = tk.Frame(dest_card, bg=COLORS['bg_primary'])
        dest_row.pack(fill="x")
        tk.Entry(dest_row, 
                textvariable=self.dest_path_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(dest_row, "浏览...", lambda: self.select_dir(self.dest_path_var)).pack(side="right")

        # 执行按钮区域
        action_frame = tk.Frame(frame, bg=COLORS['bg_primary'])
        action_frame.pack(fill="x", pady=(16, 0))
        self.create_primary_button(action_frame, "▶ 开始批量提取", self.start_extract_task).pack(fill="x")

    def start_extract_task(self):
        # Tab 1 的核心逻辑
        src_dir = self.src_path_var.get().strip()
        target_dir = self.dest_path_var.get().strip()
        name_pattern = self.name_pattern_entry.get().strip()
        ext_input = self.ext_entry.get().strip()
        use_folder_prefix = self.rename_by_folder_var.get()

        if not src_dir or not os.path.exists(src_dir):
            messagebox.showerror("错误", "源文件夹无效！")
            return
        if not target_dir:
            messagebox.showerror("错误", "目标文件夹无效！")
            return

        valid_exts = [e.strip().lower().replace('.', '') for e in ext_input.split(',') if e.strip()]
        self.log(f"--- 批量提取任务开始 ---")

        count = 0
        abs_target = os.path.abspath(target_dir)

        for root_path, dirs, files in os.walk(src_dir):
            if abs_target.startswith(os.path.abspath(root_path)): continue

            for filename in files:
                name_part, ext_part = os.path.splitext(filename)
                current_ext = ext_part.lower().replace('.', '')

                if valid_exts and (current_ext not in valid_exts): continue
                if not fnmatch.fnmatch(name_part, name_pattern) and not fnmatch.fnmatch(filename, name_pattern): continue

                src_file = os.path.join(root_path, filename)
                parent_folder_name = os.path.basename(root_path)
                dst_file = os.path.join(target_dir, filename)

                # 冲突处理
                if os.path.exists(dst_file):
                    if use_folder_prefix:
                        new_name = f"{parent_folder_name}_{filename}"
                        dst_file = os.path.join(target_dir, new_name)
                    else:
                        timestamp = datetime.now().strftime("%H%M%S_%f")[:10]
                        dst_file = os.path.join(target_dir, f"{name_part}_{timestamp}{ext_part}")

                try:
                    shutil.copy2(src_file, dst_file)
                    self.log(f"√ 提取成功: {os.path.basename(dst_file)}")
                    count += 1
                except Exception as e:
                    self.log(f"× 提取失败: {filename} - {e}")

        self.log(f"任务完成，共提取 {count} 个文件。")
        messagebox.showinfo("完成", f"成功复制 {count} 个文件。")


    # ================= Tab 2: PDF智能重命名 (Enterprise风格) =================
    def setup_tab2(self):
        self.excel_path_var = tk.StringVar()
        self.pdf_dir_var = tk.StringVar()
        self.api_key_var = tk.StringVar(value=self.config.get("api_key", ""))
        self.model_id_var = tk.StringVar(value=self.config.get("model_id", "doubao-seed-1-6-vision-250815"))
        self.ocr_mode_var = tk.StringVar(value=self.config.get("ocr_mode", "local"))
        self.df = None 
        
        # 新增变量：页码配置和文件名保留选项
        self.key_data_page_var = tk.StringVar(value="1")  # 关键数据识别页码，默认第1页
        self.title_page_var = tk.StringVar(value="1")      # PDF标题识别页码，默认第1页
        self.keep_original_filename_var = tk.BooleanVar(value=True)  # 是否保留原文件名，默认保留
        
        # 任务控制变量
        self.is_task_running = False  # 任务是否正在运行
        self.stop_task_flag = False   # 停止任务标志 

        frame = self.tab2
        frame.configure(bg=COLORS['bg_primary'])

        # 创建滚动容器
        main_canvas = tk.Canvas(frame, bg=COLORS['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=COLORS['bg_primary'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        # 使用窗口宽度让内容填满
        self.scrollable_window_tab2 = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定Canvas大小变化事件，调整内部窗口宽度
        def on_canvas_configure(event):
            canvas_width = event.width - 4  # 减去边距
            main_canvas.itemconfig(self.scrollable_window_tab2, width=canvas_width)
        main_canvas.bind('<Configure>', on_canvas_configure)
        
        # 添加鼠标滚轮支持 - 只绑定当前Tab2的Canvas
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            main_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            main_canvas.unbind_all("<MouseWheel>")
        
        # 当鼠标进入Canvas时绑定滚轮事件
        main_canvas.bind('<Enter>', bind_mousewheel)
        # 当鼠标离开Canvas时解绑滚轮事件
        main_canvas.bind('<Leave>', unbind_mousewheel)
        
        # 布局滚动区域 - 填满整个Tab，无padx
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 标题
        title_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="PDF智能重命名",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="使用OCR识别PDF内容，根据Excel数据智能重命名文件",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))

        # 1. AI配置卡片
        ai_card = self.create_card(scrollable_frame, "AI引擎配置")
        ai_card.pack(fill="x", pady=8)

        ai_grid = tk.Frame(ai_card, bg=COLORS['bg_primary'])
        ai_grid.pack(fill="x")

        # API Key
        tk.Label(ai_grid, 
                text="API Key:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(ai_grid, 
                textvariable=self.api_key_var, 
                show="●",
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        self.create_secondary_button(ai_grid, "测试连接", self.test_ai_connection).grid(row=0, column=2)

        # Model ID
        tk.Label(ai_grid, 
                text="模型ID:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Entry(ai_grid, 
                textvariable=self.model_id_var, 
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=1, column=1, padx=(12, 8), sticky="w")

        # OCR模式选择
        mode_frame = tk.Frame(ai_card, bg=COLORS['bg_primary'])
        mode_frame.pack(fill="x", pady=(12, 0))
        tk.Label(mode_frame, 
                text="识别模式:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        
        mode_options = [
            ("本地OCR (免费/快速)", "local"),
            ("火山AI (收费/高精度)", "ai")
        ]
        for text, value in mode_options:
            tk.Radiobutton(mode_frame, 
                          text=text,
                          variable=self.ocr_mode_var,
                          value=value,
                          command=self.save_config,
                          bg=COLORS['bg_primary'],
                          fg=COLORS['text_primary'],
                          selectcolor=COLORS['bg_primary'],
                          font=('微软雅黑', 9)).pack(side="left", padx=(16, 0))

        # 2. Excel配置卡片
        excel_card = self.create_card(scrollable_frame, "Excel数据源")
        excel_card.pack(fill="x", pady=8)

        # 文件选择行
        excel_row = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        excel_row.pack(fill="x", pady=(0, 12))
        tk.Entry(excel_row, 
                textvariable=self.excel_path_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(excel_row, "加载Excel", self.load_excel).pack(side="right")

        # 列选择行
        col_grid = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        col_grid.pack(fill="x")

        col_configs = [
            ("关联数据列:", "key_col_cb", 12),
            ("命名列1:", "rename1_cb", 12),
            ("命名列2 (可选):", "rename2_cb", 12),
            ("命名列3 (可选):", "rename3_cb", 12)
        ]
        
        for i, (label, attr, width) in enumerate(col_configs):
            tk.Label(col_grid, 
                    text=label,
                    bg=COLORS['bg_primary'],
                    fg=COLORS['text_secondary'],
                    font=('微软雅黑', 9)).grid(row=0, column=i*2, sticky="e", padx=(16 if i>0 else 0, 8))
            cb = ttk.Combobox(col_grid, state="readonly", width=width)
            cb.grid(row=0, column=i*2+1)
            setattr(self, attr, cb)

        # 3. PDF目录卡片
        pdf_card = self.create_card(scrollable_frame, "PDF文件夹")
        pdf_card.pack(fill="x", pady=8)

        pdf_row = tk.Frame(pdf_card, bg=COLORS['bg_primary'])
        pdf_row.pack(fill="x")
        tk.Entry(pdf_row, 
                textvariable=self.pdf_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(pdf_row, "浏览...", lambda: self.select_dir(self.pdf_dir_var)).pack(side="right")

        # 4. 识别配置卡片
        config_card = self.create_card(scrollable_frame, "识别配置")
        config_card.pack(fill="x", pady=8)

        config_grid = tk.Frame(config_card, bg=COLORS['bg_primary'])
        config_grid.pack(fill="x")

        # 页码配置 - 并排排列
        tk.Label(config_grid, 
                text="关键数据识别页码:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        key_page_entry = tk.Entry(config_grid, 
                                textvariable=self.key_data_page_var,
                                width=5,
                                bg=COLORS['input_bg'],
                                fg=COLORS['text_primary'],
                                insertbackground=COLORS['text_primary'],
                                relief='solid',
                                highlightthickness=1,
                                highlightcolor=COLORS['border'])
        key_page_entry.grid(row=0, column=1, padx=(12, 8), sticky="w")
        tk.Label(config_grid, 
                text="(第几页识别关键数据)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="w", padx=(0, 20))

        # PDF标题识别页码
        tk.Label(config_grid, 
                text="PDF标题识别页码:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=3, sticky="e", pady=6)
        title_page_entry = tk.Entry(config_grid, 
                                  textvariable=self.title_page_var,
                                  width=5,
                                  bg=COLORS['input_bg'],
                                  fg=COLORS['text_primary'],
                                  insertbackground=COLORS['text_primary'],
                                  relief='solid',
                                  highlightthickness=1,
                                  highlightcolor=COLORS['border'])
        title_page_entry.grid(row=0, column=4, padx=(12, 8), sticky="w")
        tk.Label(config_grid, 
                text="(第几页识别PDF标题)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=5, sticky="w")

        # 文件名保留选项
        filename_frame = tk.Frame(config_card, bg=COLORS['bg_primary'])
        filename_frame.pack(fill="x", pady=(12, 0))
        
        tk.Checkbutton(filename_frame, 
                      text="保留原文件名（在命名后追加原文件名）",
                      variable=self.keep_original_filename_var,
                      bg=COLORS['bg_primary'],
                      fg=COLORS['text_primary'],
                      selectcolor=COLORS['bg_primary'],
                      font=('微软雅黑', 9)).pack(anchor="w")

        # 执行按钮区域
        action_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'], pady=20)
        action_frame.pack(fill="x", side="bottom")
        
        # 按钮容器
        button_container = tk.Frame(action_frame, bg=COLORS['bg_primary'])
        button_container.pack(fill="x")
        
        self.start_button = self.create_primary_button(button_container, "▶ 开始识别并重命名", self.start_pdf_rename_task)
        self.start_button.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.stop_button = tk.Button(button_container, 
                                     text="■ 停止任务", 
                                     command=self.stop_pdf_task,
                                     bg=COLORS['danger'],
                                     fg='white',
                                     font=('微软雅黑', 10, 'bold'),
                                     relief='flat',
                                     padx=20,
                                     pady=8,
                                     cursor='hand2',
                                     state="disabled")
        self.stop_button.pack(side="right", fill="x", expand=True)

    def stop_pdf_task(self):
        """停止PDF重命名任务"""
        if self.is_task_running:
            self.stop_task_flag = True
            self.log("🛑 正在停止任务...")
            self.stop_button.configure(state="disabled", text="停止中...")
        else:
            self.log("⚠️ 当前没有正在运行的任务")

    def update_ui_state(self, is_running):
        """更新UI状态 - 任务运行时禁用控件"""
        self.is_task_running = is_running
        
        if is_running:
            # 任务运行中
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal", text="■ 停止任务")
            self.stop_task_flag = False
        else:
            # 任务停止
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled", text="■ 停止任务")

    def test_ai_connection(self):
        """测试火山引擎连通性"""
        key = self.api_key_var.get().strip()
        m_id = self.model_id_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请填写 API KEY")
            return

        self.save_config()
        self.log(f"正在测试连接: {m_id} ...")

        try:
            client = OpenAI(api_key=key, base_url="https://ark.cn-beijing.volces.com/api/v3")
            response = client.chat.completions.create(
                model=m_id,
                messages=[{"role": "user", "content": "hi"}],
                max_tokens=5
            )
            self.log("√ 连接成功！模型响应正常。")
            messagebox.showinfo("成功", "连接成功！")
        except Exception as e:
            self.log(f"× 连接失败: {e}")
            messagebox.showerror("连接失败", f"错误信息:\n{e}")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            try:
                self.excel_path_var.set(path)
                self.df = pd.read_excel(path)
                cols_required = self.df.columns.tolist()
                cols_optional = ["(不使用)"] + cols_required

                self.key_col_cb['values'] = cols_required
                self.rename1_cb['values'] = cols_required
                # 为命名列添加PDF标题选项
                rename_cols_with_title = ["(不使用)", "PDF标题"] + cols_required
                self.rename2_cb['values'] = rename_cols_with_title
                self.rename3_cb['values'] = rename_cols_with_title
                self.rename2_cb.current(0)
                self.rename3_cb.current(0)

                self.log(f"Excel 加载成功: {len(self.df)} 行")
            except Exception as e:
                messagebox.showerror("错误", f"Excel 读取失败: {e}")

    def call_volcengine_ocr(self, img_bytes, extract_title=False, key_data_hint=""):
        """调用火山引擎 Vision 模型进行 OCR"""
        try:
            client = OpenAI(
                api_key=self.api_key_var.get(),
                base_url="https://ark.cn-beijing.volces.com/api/v3"
            )
            base64_image = base64.b64encode(img_bytes).decode('utf-8')

            # 根据是否需要提取标题选择不同的提示词
            if extract_title:
                prompt_text = f"""你是一位专业的文档分析专家。请分析提供的PDF页面图像，只关注页面前200个字符区域内的内容。

【任务】从页面顶部开始的前200个字符区域内，提取以下内容：
1. 文档标题（通常字号最大、居中或加粗显示）
2. 关联数据/关键信息（能标识该文档的唯一信息，如编号、日期、名称等）

标题识别标准（按优先级）：
1. 位于页面上半部分，字体明显大于其他文本
2. 可能是加粗或居中显示
3. 长度适中（5-80个字符）
4. 不包含日期、页码、公司抬头等装饰性文字

关联数据识别：
- 寻找能唯一标识该文档的关键字段（如：编号、名称、日期、金额等）
- 返回最相关的2-3个关键信息片段

输出格式（严格遵循，不要添加额外说明）：
===TITLE_START===
[识别到的标题，如果没有则写NO_TITLE_FOUND]
===TITLE_END===

===KEY_DATA_START===
[提取的关键数据/关联信息，每行一个，最多3行]
[格式示例：
编号: XXX
日期: YYYY-MM-DD
名称: ZZZ]
===KEY_DATA_END===
"""
            else:
                if key_data_hint:
                    prompt_text = f"""请提取图片中的文字内容，特别关注与"{key_data_hint}"相关的信息。
直接输出识别到的文字，不要包含解释。"""
                else:
                    prompt_text = """请提取图片中前200个字符区域内的文字内容。
直接输出识别到的文字，不要包含解释或格式标记。"""

            response = client.chat.completions.create(
                model=self.model_id_var.get(),
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{base64_image}"
                                }
                            }
                        ]
                    }
                ]
            )
            
            result = response.choices[0].message.content
            # 获取token使用情况
            token_usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
            if hasattr(response, 'usage') and response.usage:
                token_usage = {
                    "prompt_tokens": getattr(response.usage, 'prompt_tokens', 0),
                    "completion_tokens": getattr(response.usage, 'completion_tokens', 0),
                    "total_tokens": getattr(response.usage, 'total_tokens', 0)
                }
            return result, token_usage
                
        except Exception as e:
            self.log(f"AI API 调用错误: {e}")
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

    def extract_pdf_title_ai(self, pdf_path, page_num=1):
        """使用AI OCR提取PDF标题和关联数据"""
        try:
            self.log(f"🔍 使用AI OCR识别PDF标题 (第{page_num}页，前200字符区域)...")
            doc = fitz.open(pdf_path)
            
            # 检查页码是否有效
            if page_num < 1 or page_num > len(doc):
                self.log(f"⚠️ PDF标题页码无效: {page_num} (总页数: {len(doc)})")
                doc.close()
                return ""
            
            # 获取指定页的图像
            page = doc.load_page(page_num - 1)
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            doc.close()
            
            # 调用AI OCR进行标题识别
            result, token_usage = self.call_volcengine_ocr(img_bytes, extract_title=True)
            
            # 记录token使用情况
            if token_usage['total_tokens'] > 0:
                self.log(f"📊 Token消耗: {token_usage['prompt_tokens']}输入 / {token_usage['completion_tokens']}输出 / {token_usage['total_tokens']}总计")
            
            # 使用正则表达式解析AI返回的结果
            title_match = re.search(r'===TITLE_START===(.*?)===TITLE_END===', result, re.DOTALL)
            key_data_match = re.search(r'===KEY_DATA_START===(.*?)===KEY_DATA_END===', result, re.DOTALL)
            
            title = ""
            if title_match:
                title = title_match.group(1).strip()
                if title and title != "NO_TITLE_FOUND" and len(title) > 3:
                    self.log(f"📄 AI OCR识别标题: {title}")
                else:
                    title = ""
            
            # 提取并记录关联数据
            if key_data_match:
                key_data = key_data_match.group(1).strip()
                if key_data:
                    self.log(f"📋 AI OCR提取关联数据:\n{key_data}")
            
            if not title:
                self.log(f"⚠️ AI OCR未能识别到标题")
            
            return title, token_usage
                
        except Exception as e:
            self.log(f"AI OCR标题提取错误: {e}")
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

    def _calculate_title_score_v2(self, text, y_position, font_size, font_flags, page_height, page_width, text_center_x, max_font):
        """计算标题评分V2（0-100）- 综合位置、字号、粗细、居中"""
        # 位置分：页面顶部1/3区域得满分（占20%）
        position_score = max(0, 100 - (y_position / (page_height / 3)) * 100) if y_position < page_height / 3 else 0
        
        # 字号分：相对最大字号的百分比（占35%）
        font_score = (font_size / max_font) * 100 if max_font > 0 else 0
        
        # 粗细分：检测是否为粗体（占25%）
        # PyMuPDF flags: 2^4 = 16 表示粗体 (fz_font_is_bold)
        is_bold = (font_flags & 16) != 0 or (font_flags & 2) != 0  # bold or force bold
        bold_score = 100 if is_bold else 50
        
        # 居中度分：计算文本中心与页面中心的偏移（占20%）
        page_center_x = page_width / 2
        offset = abs(text_center_x - page_center_x)
        center_score = max(0, 100 - (offset / (page_width / 6)) * 100)  # 1/6页面宽度为满分容差
        
        # 长度惩罚：过短或过长都扣分
        length = len(text)
        if 5 <= length <= 50:
            length_factor = 1.0
        else:
            length_factor = max(0.5, 1.0 - abs(length - 25) / 50)
        
        # 加权总分
        total = (position_score * 0.20 + font_score * 0.35 + bold_score * 0.25 + center_score * 0.20) * length_factor
        return min(100, total), is_bold

    def _is_likely_header_footer(self, text):
        """判断文本是否是页眉页脚等装饰性文字"""
        # 排除纯数字（页码）
        if text.isdigit():
            return True
        # 排除常见页眉格式
        if re.match(r'^\d+\s*/\s*\d+$', text):  # "1 / 10" 页码格式
            return True
        if re.match(r'^(第\s*\d+\s*页|Page\s*\d+).*$', text, re.IGNORECASE):
            return True
        # 排除过短的文本
        if len(text) < 5:
            return True
        # 排除常见公司抬头
        if any(keyword in text for keyword in ['有限公司', '股份有限公司', '集团', 'Corp', 'Inc', 'Ltd']):
            if len(text) < 20:  # 短的公司名可能是抬头
                return True
        return False

    def _extract_title_from_text_pdf(self, page, page_num):
        """从文本型PDF中提取标题 - 基于PyMuPDF（字体大小、粗细、居中）"""
        page_rect = page.rect
        page_height = page_rect.height
        page_width = page_rect.width
        
        # 获取所有文本块
        blocks = page.get_text("dict")["blocks"]
        
        candidates = []
        max_font = 0
        
        # 第一遍：找出最大字号和统计信息
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    if "spans" in line:
                        for span in line["spans"]:
                            max_font = max(max_font, span.get("size", 0))
        
        self.log(f"📊 页面最大字号: {max_font:.1f}pt, 页面尺寸: {page_width:.0f}x{page_height:.0f}")
        
        # 第二遍：收集候选标题
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    if "spans" in line:
                        # 合并同一行的所有span
                        text_parts = []
                        line_font = 0
                        line_flags = 0
                        line_bbox = line.get("bbox", [0, 0, 0, 0])
                        
                        for span in line["spans"]:
                            text_parts.append(span.get("text", ""))
                            line_font = max(line_font, span.get("size", 0))
                            line_flags = max(line_flags, span.get("flags", 0))
                        
                        text = "".join(text_parts).strip()
                        
                        # 基本过滤
                        if not text or len(text) < 3 or len(text) > 100:
                            continue
                        
                        # 排除装饰性文字
                        if self._is_likely_header_footer(text):
                            continue
                        
                        # 计算位置（使用文本块中心Y坐标）
                        y_pos = (line_bbox[1] + line_bbox[3]) / 2
                        
                        # 计算文本中心X坐标
                        text_center_x = (line_bbox[0] + line_bbox[2]) / 2
                        
                        # 计算评分
                        score, is_bold = self._calculate_title_score_v2(
                            text, y_pos, line_font, line_flags, 
                            page_height, page_width, text_center_x, max_font
                        )
                        
                        # 只保留得分高的候选（>35分）
                        if score > 35:
                            candidates.append({
                                "text": text,
                                "score": score,
                                "font": line_font,
                                "is_bold": is_bold,
                                "center_offset": abs(text_center_x - page_width/2),
                                "y": y_pos
                            })
        
        # 按得分排序
        candidates.sort(key=lambda x: x["score"], reverse=True)
        
        if candidates:
            best = candidates[0]
            bold_str = "粗体" if best['is_bold'] else "常规"
            self.log(f"🔍 PyMuPDF发现候选标题: \"{best['text']}\" (得分: {best['score']:.1f}, 字号: {best['font']:.1f}pt, {bold_str}, 偏移: {best['center_offset']:.0f}px)")
            
            # 输出前3个候选供调试
            for i, c in enumerate(candidates[:3], 1):
                if i > 1:
                    b_str = "粗体" if c['is_bold'] else "常规"
                    self.log(f"   候选{i}: \"{c['text']}\" (得分: {c['score']:.1f}, {c['font']:.1f}pt, {b_str})")
            
            return best["text"]
        
        return None

    def _extract_title_from_scanned_pdf(self, img_bytes, page_num):
        """从扫描型PDF中提取标题 - 使用RapidOCR"""
        if self.ocr_engine is None:
            self.log("初始化本地 OCR 引擎...")
            self.ocr_engine = RapidOCR()
        
        result, _ = self.ocr_engine(img_bytes)
        if not result:
            return None
        
        # 解析OCR结果，分析位置和文字
        candidates = []
        for line_data in result:
            # line_data格式: [bbox, text, confidence]
            if len(line_data) >= 2:
                bbox = line_data[0]  # [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                text = line_data[1].strip()
                
                if not text or len(text) < 3 or len(text) > 100:
                    continue
                
                # 排除装饰性文字
                if self._is_likely_header_footer(text):
                    continue
                
                # 计算Y位置（顶部Y坐标越小，位置越靠上）
                y_pos = min(p[1] for p in bbox)
                
                # 扫描型PDF无法获取字号，主要依据位置和长度
                length = len(text)
                if 10 <= length <= 50:
                    score = 70  # 基础分
                    # 位置越靠上，分数越高（假设页面高度1000px）
                    score += max(0, 30 - y_pos / 10)
                    
                    candidates.append({
                        "text": text,
                        "score": score,
                        "y": y_pos
                    })
        
        # 按得分排序
        candidates.sort(key=lambda x: x["score"], reverse=True)
        
        if candidates:
            best = candidates[0]
            self.log(f"🔍 OCR发现候选标题: \"{best['text']}\" (得分: {best['score']:.1f})")
            return best["text"]
        
        return None

    def extract_pdf_title(self, pdf_path, page_num=1):
        """提取PDF标题 - 多策略综合识别，返回 (title, token_usage)"""
        token_usage_total = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
        try:
            # 如果使用AI OCR，优先使用AI OCR
            if self.ocr_mode_var.get() == "ai":
                title, token_usage = self.extract_pdf_title_ai(pdf_path, page_num)
                return title, token_usage
            
            # 本地OCR模式：完全基于页面字体特征识别（不使用PDF元数据）
            self.log(f"🔍 开始识别PDF标题 (第{page_num}页) - 本地OCR模式...")
            
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                doc.close()
                return "", token_usage_total
            
            # 检查页码是否有效
            if page_num < 1 or page_num > len(doc):
                self.log(f"⚠️ PDF标题页码无效: {page_num} (总页数: {len(doc)})")
                doc.close()
                return "", token_usage_total
            
            # 加载指定页
            page = doc.load_page(page_num - 1)
            
            # 策略1: 尝试从文本型PDF提取（PyMuPDF - 基于字体大小、粗细、居中）
            text_title = self._extract_title_from_text_pdf(page, page_num)
            if text_title:
                self.log(f"📄 PyMuPDF识别标题: {text_title}")
                doc.close()
                return text_title, token_usage_total
            
            # 策略2: 扫描型PDF - 使用OCR
            self.log("🔍 尝试使用OCR识别扫描型PDF...")
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            
            ocr_title = self._extract_title_from_scanned_pdf(img_bytes, page_num)
            if ocr_title:
                self.log(f"📄 OCR识别标题: {ocr_title}")
                doc.close()
                return ocr_title, token_usage_total
            
            self.log(f"⚠️ 未能识别到PDF标题")
            doc.close()
            return "", token_usage_total
            
        except Exception as e:
            self.log(f"PDF标题提取错误: {e}")
            return "", token_usage_total

    def get_pdf_text_content(self, pdf_path, page_num=1):
        """获取指定页的文本内容用于关键数据识别"""
        text_content = ""
        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                doc.close()
                return ""
            
            # 检查页码是否有效
            if page_num < 1 or page_num > len(doc):
                self.log(f"⚠️ 关键数据页码无效: {page_num} (总页数: {len(doc)})")
                doc.close()
                return ""
            
            # 只处理指定页
            page = doc.load_page(page_num - 1)  # 转换为0基索引
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")

            if self.ocr_mode_var.get() == "local":
                if self.ocr_engine is None: 
                    self.log("初始化本地 OCR 引擎...")
                    self.ocr_engine = RapidOCR()
                result, _ = self.ocr_engine(img_bytes)
                if result: 
                    text_content = "".join([line[1] for line in result])
                    self.log(f"  本地OCR提取内容长度: {len(text_content)} 字符")
                else:
                    self.log(f"  ⚠️ 本地OCR未识别到任何文本")
                return text_content, {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
            else:
                self.log(f"正在请求 AI 识别: {os.path.basename(pdf_path)} (第{page_num}页)...")
                text_content, token_usage = self.call_volcengine_ocr(img_bytes)
                if token_usage['total_tokens'] > 0:
                    self.log(f"📊 Token消耗: {token_usage['prompt_tokens']}输入 / {token_usage['completion_tokens']}输出 / {token_usage['total_tokens']}总计")
                return text_content, token_usage
        except Exception as e:
            self.log(f"文件读取错误: {e}")
            return ""

    def _sanitize_filename(self, filename):
        """清理文件名中的非法字符，确保Windows文件名合法"""
        # Windows非法字符: < > : " / \ | ? *
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '')
        # 移除控制字符 (0-31)
        filename = ''.join(char for char in filename if ord(char) >= 32)
        # 移除首尾空格和点
        filename = filename.strip(' .')
        # 如果文件名为空，返回默认名
        if not filename:
            filename = "unnamed"
        return filename

    def _normalize_for_matching(self, text):
        """标准化文本用于匹配 - 移除空格、统一大小写"""
        if not text:
            return ""
        # 移除所有空白字符（空格、制表符、换行等）
        return ''.join(text.split()).lower()

    def _find_best_match(self, content, data_map):
        """模糊匹配：处理OCR引入的空格差异"""
        content_normalized = self._normalize_for_matching(content)
        
        # 第一遍：精确匹配
        for key in data_map.keys():
            if key in content:
                return key
        
        # 第二遍：标准化匹配（忽略空格差异）
        for key in data_map.keys():
            key_normalized = self._normalize_for_matching(key)
            if key_normalized and key_normalized in content_normalized:
                # 记录匹配详情便于调试
                self.log(f"🔍 模糊匹配成功: \"{key}\" (标准化: {key_normalized})")
                return key
        
        return None

    def start_pdf_rename_task(self):
        """启动PDF重命名任务 - 使用线程防止界面假死"""
        # 检查配置
        if not self.df is not None:
            messagebox.showwarning("配置不全", "请先加载Excel文件")
            return
        
        key_col = self.key_col_cb.get()
        r1_col = self.rename1_cb.get()
        if not key_col or not r1_col:
            messagebox.showwarning("配置不全", "请确保选择了关联列和至少一个命名列")
            return
        
        # 获取页码配置
        try:
            key_data_page = int(self.key_data_page_var.get().strip())
            title_page = int(self.title_page_var.get().strip())
        except ValueError:
            messagebox.showerror("配置错误", "页码必须是数字")
            return
        
        # 检查是否需要PDF标题识别
        r2_col = self.rename2_cb.get()
        r3_col = self.rename3_cb.get()
        need_title = (r2_col == "PDF标题" or r3_col == "PDF标题")
        
        # 启动线程执行任务
        thread = threading.Thread(
            target=self._run_pdf_rename_task,
            args=(key_col, r1_col, r2_col, r3_col, key_data_page, title_page),
            daemon=True
        )
        thread.start()

    def _run_pdf_rename_task(self, key_col, r1_col, r2_col, r3_col, key_data_page, title_page):
        """在线程中执行PDF重命名任务"""
        try:
            self.update_ui_state(True)
            
            pdf_dir = self.pdf_dir_var.get()
            keep_original = self.keep_original_filename_var.get()

            self.log(f"--- 开始重命名任务 (模式: {self.ocr_mode_var.get()}) ---")
            self.log(f"关键数据识别页码: 第{key_data_page}页")
            self.log(f"PDF标题识别页码: 第{title_page}页")
            self.log(f"文件名处理: {'保留原文件名' if keep_original else '完全替换'}")

            # 构建数据映射
            data_map = {}
            for _, row in self.df.iterrows():
                if self.stop_task_flag:
                    break
                    
                k = str(row[key_col]).strip()
                if not k or k == 'nan': 
                    continue

                # 获取命名列1的值
                part1 = str(row[r1_col]).strip()
                if not part1 or part1 == 'nan':
                    continue

                # 构建带位置信息的文件名组件列表
                # 每个组件: (值, 是否为PDF标题, 原始列位置)
                name_parts = [(part1, False, 1)]

                # 处理命名列2
                if r2_col and r2_col != "(不使用)":
                    if r2_col == "PDF标题":
                        name_parts.append(("", True, 2))  # 标记为PDF标题，稍后填充
                    else:
                        part2 = str(row[r2_col]).strip()
                        if part2 and part2 != 'nan':
                            name_parts.append((part2, False, 2))

                # 处理命名列3
                if r3_col and r3_col != "(不使用)":
                    if r3_col == "PDF标题":
                        name_parts.append(("", True, 3))  # 标记为PDF标题，稍后填充
                    else:
                        part3 = str(row[r3_col]).strip()
                        if part3 and part3 != 'nan':
                            name_parts.append((part3, False, 3))

                data_map[k] = {
                    'parts': name_parts,
                    'use_pdf_title': r2_col == "PDF标题" or r3_col == "PDF标题"
                }

            # 处理PDF文件
            pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
            success_count = 0
            total = len(pdf_files)

            # 初始化token使用统计
            token_usage_total = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

            for i, filename in enumerate(pdf_files):
                if self.stop_task_flag:
                    self.log("🛑 任务已被用户停止")
                    break
                    
                # 更新UI显示进度
                self.root.after(0, lambda: None)  # 保持UI响应
                
                pdf_path = os.path.join(pdf_dir, filename)
                
                # 使用指定页码获取关键数据
                content, content_tokens = self.get_pdf_text_content(pdf_path, key_data_page)
                # 累加token使用
                for key in token_usage_total:
                    token_usage_total[key] += content_tokens.get(key, 0)
                
                # 只有在需要时才提取PDF标题
                pdf_title = ""
                title_tokens = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
                if data_map and any(data_map[k]['use_pdf_title'] for k in data_map.keys()):
                    pdf_title, title_tokens = self.extract_pdf_title(pdf_path, title_page)
                    # 累加token使用
                    for key in token_usage_total:
                        token_usage_total[key] += title_tokens.get(key, 0)
                    # 清理标题中的非法字符
                    if pdf_title:
                        pdf_title = self._sanitize_filename(pdf_title)
                        self.log(f"📄 清理后的PDF标题: {pdf_title}")

                matched_key = self._find_best_match(content, data_map)

                if not matched_key:
                    self.log(f"  ⚠️ 未匹配: 内容长度 {len(content)} 字符, 前100字符: {content[:100]}...")
                    self.log(f"? 未匹配: {filename}")
                    continue

                # 匹配成功，按正确顺序构建文件名
                parts_data = data_map[matched_key]['parts']
                final_parts = []
                
                for value, is_pdf_title, position in parts_data:
                    if is_pdf_title:
                        if pdf_title:
                            final_parts.append(pdf_title)
                    else:
                        # 清理普通列的值
                        sanitized = self._sanitize_filename(value)
                        if sanitized:
                            final_parts.append(sanitized)
                
                # 用下划线连接各部分
                new_prefix = "_".join(final_parts) if final_parts else "unnamed"

                # 根据选项决定文件名格式
                if keep_original:
                    new_filename = f"{new_prefix}_{filename}"
                else:
                    # 完全替换原文件名，保留扩展名
                    name_part, ext_part = os.path.splitext(filename)
                    new_filename = f"{new_prefix}{ext_part}"
                
                # 最终清理确保文件名合法
                new_filename = self._sanitize_filename(new_filename)

                new_path = os.path.join(pdf_dir, new_filename)

                try:
                    if not os.path.exists(new_path):
                        os.rename(pdf_path, new_path)
                        self.log(f"√ 重命名: {new_filename}")
                        success_count += 1
                    else:
                        self.log(f"! 跳过(已存在): {new_filename}")
                except Exception as e:
                    self.log(f"× 错误: {e}")

            if not self.stop_task_flag:
                # 打印Token消耗汇总
                if token_usage_total['total_tokens'] > 0:
                    self.log(f"📊📊📊 本次任务Token消耗汇总 📊📊📊")
                    self.log(f"📊 输入Token: {token_usage_total['prompt_tokens']}")
                    self.log(f"📊 输出Token: {token_usage_total['completion_tokens']}")
                    self.log(f"📊 总计Token: {token_usage_total['total_tokens']}")
                self.log(f"任务完成，成功处理: {success_count} / {total}")
                self.root.after(0, lambda: messagebox.showinfo("完成", f"任务结束\n成功处理: {success_count} / {total}"))
            else:
                # 打印Token消耗汇总（即使任务被停止）
                if token_usage_total['total_tokens'] > 0:
                    self.log(f"📊📊📊 本次任务Token消耗汇总（任务已停止） 📊📊📊")
                    self.log(f"📊 输入Token: {token_usage_total['prompt_tokens']}")
                    self.log(f"📊 输出Token: {token_usage_total['completion_tokens']}")
                    self.log(f"📊 总计Token: {token_usage_total['total_tokens']}")
                self.log(f"任务已停止，已处理: {success_count} / {total}")
                
        except Exception as e:
            self.log(f"任务执行错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"任务执行失败:\n{e}"))
        finally:
            self.update_ui_state(False)

    # ================= Tab 3: 高级文件整理 (Enterprise风格) =================
    def setup_tab3(self):
        # Tab 3 variables
        self.excel_file_var = tk.StringVar()
        self.dest_root_var = tk.StringVar()
        self.file_mode_var = tk.StringVar(value="copy")  # copy or cut
        self.df_tab3 = None  # Excel data for Tab 3
        self.root_dirs_list = []  # List of source directories
        
        # 新的变量：目录结构配置
        self.directory_structure = []  # 存储目录结构 [{level: int, excel_column: str, keywords: []}]
        self.association_column = tk.StringVar()  # 关联数据列
        
        frame = self.tab3
        frame.configure(bg=COLORS['bg_primary'])
        
        # 标题区域
        title_frame = tk.Frame(frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="高级文件整理",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="基于Excel配置，按规则自动整理文件到多级目录",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))

        # 创建主滚动区域 - 填满整个宽度，无右侧空白
        main_canvas = tk.Canvas(frame, bg=COLORS['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=COLORS['bg_primary'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        # 使用窗口宽度让内容填满
        self.scrollable_window = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定Canvas大小变化事件，调整内部窗口宽度
        def on_canvas_configure(event):
            canvas_width = event.width - 4  # 减去边距
            main_canvas.itemconfig(self.scrollable_window, width=canvas_width)
        main_canvas.bind('<Configure>', on_canvas_configure)
        
        # 添加鼠标滚轮支持
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 布局滚动区域 - 填满整个Tab，无padx
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 1. 源目录配置卡片
        root_card = self.create_card(scrollable_frame, "源目录配置")
        root_card.pack(fill="x", pady=8)
        
        self.dir_list_frame = tk.Frame(root_card, bg=COLORS['bg_primary'])
        self.dir_list_frame.pack(fill="x", pady=(0, 8))
        
        btn_frame = tk.Frame(root_card, bg=COLORS['bg_primary'])
        btn_frame.pack(fill="x")
        
        self.add_dir_btn = self.create_secondary_button(btn_frame, "+ 添加目录", self.add_source_directory)
        self.add_dir_btn.pack(side="left", padx=(0, 8))
        
        self.remove_dir_btn = self.create_secondary_button(btn_frame, "- 删除选中", self.remove_source_directory)
        self.remove_dir_btn.configure(state="disabled")
        self.remove_dir_btn.pack(side="left")
        
        tk.Label(btn_frame, 
               text="(最多10个目录)", 
               bg=COLORS['bg_primary'],
               fg=COLORS['text_secondary'],
               font=('微软雅黑', 9)).pack(side="left", padx=(16, 0))
        
        # 2. Excel配置卡片
        excel_card = self.create_card(scrollable_frame, "Excel配置")
        excel_card.pack(fill="x", pady=8)
        
        # Excel文件行
        excel_row = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        excel_row.pack(fill="x", pady=(0, 12))
        tk.Entry(excel_row, 
                textvariable=self.excel_file_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(excel_row, "选择Excel", self.load_excel_tab3).pack(side="right")
        
        # 关联数据列
        assoc_frame = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        assoc_frame.pack(fill="x")
        
        tk.Label(assoc_frame, 
                text="关联数据列:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        self.association_col_cb = ttk.Combobox(assoc_frame, state="readonly", width=20)
        self.association_col_cb.pack(side="left", padx=(12, 0))
        
        # 3. 目录结构编辑器卡片
        structure_card = self.create_card(scrollable_frame, "目录结构编辑器")
        structure_card.pack(fill="both", expand=True, pady=8)
        
        # 工具栏
        toolbar_frame = tk.Frame(structure_card, bg=COLORS['bg_primary'])
        toolbar_frame.pack(fill="x", pady=(0, 8))
        
        self.create_secondary_button(toolbar_frame, "+ 添加目录层级", self.add_directory_level).pack(side="left", padx=(0, 8))
        self.create_secondary_button(toolbar_frame, "- 删除选中", self.delete_selected_tree_item).pack(side="left", padx=(0, 8))
        self.create_secondary_button(toolbar_frame, "+ 添加关键字", self.add_keyword_to_selected).pack(side="left")
        
        tk.Label(toolbar_frame, 
               text="(最多5级目录)", 
               bg=COLORS['bg_primary'],
               fg=COLORS['text_secondary'],
               font=('微软雅黑', 9)).pack(side="left", padx=(16, 0))
        
        # CheckboxTreeview 编辑区域
        try:
            from ttkwidgets import CheckboxTreeview
            
            tree_container = tk.Frame(structure_card, bg=COLORS['bg_primary'])
            tree_container.pack(fill="both", expand=True)
            
            self.tree = CheckboxTreeview(tree_container, 
                                        columns=("excel_column", "type"), 
                                        displaycolumns=("excel_column",),
                                        height=12)
            self.tree.heading("#0", text="目录结构")
            self.tree.heading("excel_column", text="Excel列")
            self.tree.column("#0", width=280)
            self.tree.column("excel_column", width=140)
            
            # 配置树形样式 - Enterprise风格
            self.tree.tag_configure("directory", 
                                   background=COLORS['bg_secondary'],
                                   foreground=COLORS['text_primary'])
            self.tree.tag_configure("keyword", 
                                   background=COLORS['bg_primary'],
                                   foreground=COLORS['text_secondary'])
            
            # 滚动条
            tree_scroll = tk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=tree_scroll.set)
            
            self.tree.pack(side="left", fill="both", expand=True)
            tree_scroll.pack(side="right", fill="y")
            
            # 绑定事件
            self.tree.bind("<Double-1>", self.on_tree_double_click)
            self.tree.bind("<Button-3>", self.on_tree_right_click)
            
        except ImportError:
            error_label = tk.Label(structure_card, 
                                 text="请安装 ttkwidgets: pip install ttkwidgets", 
                                 fg=COLORS['danger'],
                                 bg=COLORS['bg_primary'],
                                 font=('微软雅黑', 10))
            error_label.pack(pady=20)
        
        # 4. 操作设置卡片
        settings_card = self.create_card(scrollable_frame, "操作设置")
        settings_card.pack(fill="x", pady=8)
        
        # 文件操作模式
        mode_frame = tk.Frame(settings_card, bg=COLORS['bg_primary'])
        mode_frame.pack(fill="x", pady=(0, 12))
        
        tk.Label(mode_frame, 
                text="文件操作:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        
        mode_options = [
            ("复制文件 (保留原文件)", "copy"),
            ("剪切文件 (移动文件)", "cut")
        ]
        
        for text, value in mode_options:
            tk.Radiobutton(mode_frame, 
                          text=text,
                          variable=self.file_mode_var,
                          value=value,
                          bg=COLORS['bg_primary'],
                          fg=COLORS['text_primary'],
                          selectcolor=COLORS['bg_primary'],
                          font=('微软雅黑', 9)).pack(side="left", padx=(16, 0))
        
        # 目标目录
        dest_row = tk.Frame(settings_card, bg=COLORS['bg_primary'])
        dest_row.pack(fill="x")
        
        tk.Label(dest_row, 
                text="目标目录:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(dest_row, 
                textvariable=self.dest_root_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8), fill="x", expand=True)
        self.create_secondary_button(dest_row, "浏览...", lambda: self.select_dir(self.dest_root_var)).pack(side="right")
        
        # 执行按钮区域
        action_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'])
        action_frame.pack(fill="x", pady=(16, 0))
        self.create_primary_button(action_frame, "▶ 开始高级整理", self.start_advanced_organize_task).pack(fill="x")
        
        # 初始化
        self.update_tree_display()

    # ============ Tab 3 新增功能函数 ============
    def update_tree_display(self):
        """更新CheckboxTreeview显示"""
        if not hasattr(self, 'tree'):
            return
        
        # 清空现有内容
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.directory_structure:
            # 显示提示信息
            self.tree.insert("", "end", text="请添加目录层级...", values=("",), tags=("placeholder",))
            return
        
        # 构建树形结构
        parent_map = {}  # 存储父节点映射
        
        for i, level_data in enumerate(self.directory_structure):
            level = level_data['level']
            excel_column = level_data['excel_column']
            keywords = level_data['keywords']
            
            # 创建目录节点
            if level == 0:
                parent = ""
            else:
                parent = parent_map.get(level - 1, "")
            
            dir_node = self.tree.insert(
                parent, "end", 
                text=f"第{level+1}级目录", 
                values=(excel_column,),
                tags=("directory",)
            )
            parent_map[level] = dir_node
            
            # 添加关键字节点
            for keyword in keywords:
                self.tree.insert(
                    dir_node, "end",
                    text=keyword,
                    values=("",),
                    tags=("keyword",)
                )
        
        # 展开所有节点
        self.tree.expand_all()
    
    def add_directory_level(self):
        """添加目录层级"""
        if len(self.directory_structure) >= 5:
            messagebox.showwarning("限制", "最多只能添加5级目录")
            return
        
        if not self.df_tab3 is not None:
            messagebox.showwarning("提示", "请先加载Excel文件")
            return
        
        # 创建居中模态对话框
        dialog, close_dialog, overlay = self.create_centered_dialog("添加目录层级", 300, 150)
        
        # 内容框架
        content_frame = tk.Frame(dialog, bg=COLORS['bg_tertiary'])
        content_frame.pack(fill="both", expand=True, padx=20, pady=15)
        
        tk.Label(content_frame, 
                text="选择Excel列作为目录层级:", 
                bg=COLORS['bg_tertiary'], 
                fg=COLORS['text_primary'],
                font=('微软雅黑', 10)).pack(pady=(0, 10))
        
        col_cb = ttk.Combobox(content_frame, state="readonly", width=25)
        col_cb['values'] = list(self.df_tab3.columns)
        col_cb.pack(pady=5)
        
        def confirm():
            selected_col = col_cb.get()
            if selected_col:
                level_data = {
                    'level': len(self.directory_structure),
                    'excel_column': selected_col,
                    'keywords': []
                }
                self.directory_structure.append(level_data)
                self.update_tree_display()
                close_dialog()
        
        tk.Button(content_frame, 
                 text="确定", 
                 command=confirm, 
                 bg=COLORS['primary'], 
                 fg='white',
                 font=('微软雅黑', 9),
                 relief='flat',
                 padx=20,
                 pady=4).pack(pady=(15, 0))
    
    def delete_selected_tree_item(self):
        """删除选中的树形项目"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择要删除的项目")
            return
        
        item = selection[0]
        item_tags = self.tree.item(item, "tags")
        
        if "directory" in item_tags:
            # 获取目录层级
            item_text = self.tree.item(item, "text")
            if "第1级目录" in item_text:
                messagebox.showwarning("限制", "第一级目录不允许删除")
                return
            
            # 找到对应的层级索引
            level_index = None
            for i, level_data in enumerate(self.directory_structure):
                if f"第{i+1}级目录" == item_text:
                    level_index = i
                    break
            
            if level_index is not None:
                del self.directory_structure[level_index]
                # 重新编号
                for i, level_data in enumerate(self.directory_structure):
                    level_data['level'] = i
                self.update_tree_display()
        
        elif "keyword" in item_tags:
            # 删除关键字
            parent = self.tree.parent(item)
            parent_text = self.tree.item(parent, "text")
            keyword_text = self.tree.item(item, "text")
            
            # 找到对应的目录层级
            level_index = None
            for i, level_data in enumerate(self.directory_structure):
                if f"第{i+1}级目录" == parent_text:
                    level_index = i
                    break
            
            if level_index is not None and keyword_text in self.directory_structure[level_index]['keywords']:
                self.directory_structure[level_index]['keywords'].remove(keyword_text)
                self.update_tree_display()
    
    def add_keyword_to_selected(self):
        """向选中的目录添加关键字"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个目录层级")
            return
        
        item = selection[0]
        item_tags = self.tree.item(item, "tags")
        
        if "keyword" in item_tags:
            # 如果选中的是关键字，则向其父目录添加
            item = self.tree.parent(item)
            item_tags = self.tree.item(item, "tags")
        
        if "directory" in item_tags:
            # 创建居中模态对话框
            dialog, close_dialog, overlay = self.create_centered_dialog("添加关键字", 300, 140)
            
            # 内容框架
            content_frame = tk.Frame(dialog, bg=COLORS['bg_tertiary'])
            content_frame.pack(fill="both", expand=True, padx=20, pady=15)
            
            tk.Label(content_frame, 
                    text="请输入关键字:", 
                    bg=COLORS['bg_tertiary'], 
                    fg=COLORS['text_primary'],
                    font=('微软雅黑', 10)).pack(pady=(0, 10))
            
            entry = tk.Entry(content_frame, 
                           width=30, 
                           bg=COLORS['input_bg'], 
                           fg=COLORS['text_primary'], 
                           insertbackground=COLORS['text_primary'],
                           relief='solid',
                           highlightthickness=1,
                           highlightcolor=COLORS['border'])
            entry.pack(pady=5)
            entry.focus()
            
            def confirm():
                keyword = entry.get().strip()
                if keyword:
                    # 找到对应的目录层级
                    item_text = self.tree.item(item, "text")
                    level_index = None
                    for i, level_data in enumerate(self.directory_structure):
                        if f"第{i+1}级目录" == item_text:
                            level_index = i
                            break
                    
                    if level_index is not None:
                        keywords = self.directory_structure[level_index]['keywords']
                        if keyword not in keywords:
                            keywords.append(keyword)
                            self.update_tree_display()
                            close_dialog()
                        else:
                            messagebox.showwarning("警告", "关键字已存在")
            
            tk.Button(content_frame, 
                     text="确定", 
                     command=confirm, 
                     bg=COLORS['primary'], 
                     fg='white',
                     font=('微软雅黑', 9),
                     relief='flat',
                     padx=20,
                     pady=4).pack(pady=(15, 0))
            
            entry.bind('<Return>', lambda e: confirm())
    
    def on_tree_double_click(self, event):
        """处理树形控件双击事件"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_tags = self.tree.item(item, "tags")
        
        if "directory" in item_tags and self.df_tab3 is not None:
            # 编辑Excel列
            current_value = self.tree.item(item, "values")[0]
            
            # 创建居中模态对话框
            dialog, close_dialog, overlay = self.create_centered_dialog("选择Excel列", 250, 140)
            
            # 内容框架
            content_frame = tk.Frame(dialog, bg=COLORS['bg_tertiary'])
            content_frame.pack(fill="both", expand=True, padx=20, pady=15)
            
            tk.Label(content_frame, 
                    text="选择Excel列:", 
                    bg=COLORS['bg_tertiary'], 
                    fg=COLORS['text_primary'],
                    font=('微软雅黑', 10)).pack(pady=(0, 10))
            
            col_cb = ttk.Combobox(content_frame, state="readonly", width=20)
            col_cb['values'] = list(self.df_tab3.columns)
            col_cb.set(current_value)
            col_cb.pack(pady=5)
            
            def confirm():
                selected_col = col_cb.get()
                if selected_col:
                    # 找到对应的目录层级
                    item_text = self.tree.item(item, "text")
                    level_index = None
                    for i, level_data in enumerate(self.directory_structure):
                        if f"第{i+1}级目录" == item_text:
                            level_index = i
                            break
                    
                    if level_index is not None:
                        self.directory_structure[level_index]['excel_column'] = selected_col
                        self.update_tree_display()
                        close_dialog()
            
            tk.Button(content_frame, 
                     text="确定", 
                     command=confirm, 
                     bg=COLORS['primary'], 
                     fg='white',
                     font=('微软雅黑', 9),
                     relief='flat',
                     padx=20,
                     pady=4).pack(pady=(15, 0))
    
    def on_tree_right_click(self, event):
        """处理右键菜单"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_tags = self.tree.item(item, "tags")
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        
        if "directory" in item_tags:
            context_menu.add_command(label="编辑Excel列", command=lambda: self.on_tree_double_click(None))
            context_menu.add_command(label="添加关键字", command=self.add_keyword_to_selected)
            
            item_text = self.tree.item(item, "text")
            if "第1级目录" not in item_text:
                context_menu.add_separator()
                context_menu.add_command(label="删除目录", command=self.delete_selected_tree_item)
        
        elif "keyword" in item_tags:
            context_menu.add_command(label="删除关键字", command=self.delete_selected_tree_item)
        
        # 显示菜单
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def get_all_keywords(self):
        """获取所有层级的关键字"""
        all_keywords = []
        for level_data in self.directory_structure:
            all_keywords.extend(level_data['keywords'])
        return list(set(all_keywords))  # 去重
    
    def get_directory_columns(self):
        """获取目录列配置"""
        return [level_data['excel_column'] for level_data in self.directory_structure]

    def add_source_directory(self):
        """添加源目录"""
        if len(self.root_dirs_list) >= 10:
            messagebox.showwarning("限制", "最多只能添加10个源目录")
            return
        
        directory = filedialog.askdirectory()
        if directory and directory not in self.root_dirs_list:
            self.root_dirs_list.append(directory)
            self.update_directory_list()
            self.log(f"添加源目录: {directory}")

    def remove_source_directory(self):
        """删除选中的源目录"""
        selection = self.dir_listbox.curselection()
        if selection:
            index = selection[0]
            removed_dir = self.root_dirs_list.pop(index)
            self.update_directory_list()
            self.log(f"删除源目录: {removed_dir}")

    def update_directory_list(self):
        """更新目录列表显示"""
        # Clear existing widgets
        for widget in self.dir_list_frame.winfo_children():
            widget.destroy()
        
        # Create listbox with scrollbar
        list_frame = tk.Frame(self.dir_list_frame)
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.dir_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=4)
        self.dir_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.dir_listbox.yview)
        
        # Add directories to listbox
        for dir_path in self.root_dirs_list:
            self.dir_listbox.insert(tk.END, dir_path)
        
        # Update button states
        self.add_dir_btn.config(state="normal" if len(self.root_dirs_list) < 10 else "disabled")
        self.remove_dir_btn.config(state="normal" if len(self.root_dirs_list) > 0 else "disabled")

    def load_excel_tab3(self):
        """加载Excel文件并更新列选择器"""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            try:
                self.excel_file_var.set(path)
                self.df_tab3 = pd.read_excel(path)
                
                # 更新关联数据列下拉框
                cols = self.df_tab3.columns.tolist()
                self.association_col_cb['values'] = cols
                
                # 清空现有目录列配置
                self.directory_columns = []
                self.update_directory_columns_display()
                
                self.log(f"Excel 加载成功: {len(self.df_tab3)} 行, {len(cols)} 列")
            except Exception as e:
                messagebox.showerror("错误", f"Excel 读取失败: {e}")
                self.df_tab3 = None

    def start_organize_task(self):
        """执行文件整理任务"""
        # 获取配置
        if not self.root_dirs_list:
            messagebox.showerror("错误", "请至少添加一个源目录")
            return
        
        root_dirs = self.root_dirs_list.copy()
        excel_file = self.excel_file_var.get().strip()
        col_debtor = self.debtor_col_cb.get()
        col_province = self.province_col_cb.get()
        keywords_text = self.keywords_var.get().strip()
        dest_root = self.dest_root_var.get().strip()
        file_mode = self.file_mode_var.get()
        
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("错误", "Excel文件路径无效")
            return
        
        if not col_debtor or not col_province:
            messagebox.showerror("错误", "请选择Excel列名")
            return
        
        if not dest_root:
            messagebox.showerror("错误", "请选择目标目录")
            return
        
        keywords = [kw.strip() for kw in keywords_text.split(',') if kw.strip()]
        
        self.log(f"--- 开始文件整理任务 ({'剪切模式' if file_mode == 'cut' else '复制模式'}) ---")
        self.log(f"源目录: {len(root_dirs)} 个")
        self.log(f"Excel文件: {excel_file}")
        self.log(f"关键词: {len(keywords)} 个")
        
        try:
            # 执行整理逻辑
            person_province = self.load_person_province_map(excel_file, col_debtor, col_province)
            persons = set(person_province.keys())
            self.log(f"从Excel提取 {len(persons)} 条人名+省份记录")
            
            all_files = self.collect_all_files(root_dirs)
            self.log(f"共发现 {len(all_files)} 个原始文件")
            
            grouped = self.group_files_by_person(all_files, persons, keywords)
            latest = self.select_latest(grouped)
            
            self.copy_and_report(latest, person_province, Path(dest_root), keywords, file_mode)
            self.log("整理完成！")
            messagebox.showinfo("完成", f"文件整理任务完成！({'剪切' if file_mode == 'cut' else '复制'}模式)")
            
        except Exception as e:
            self.log(f"整理失败: {e}")
            messagebox.showerror("错误", f"整理失败: {e}")

    # ============ Tab 3 核心功能函数 ============
    def load_person_province_map(self, excel_path: str, col_name: str, col_prov: str) -> Dict[str, str]:
        """返回 {人名: 省份} 字典"""
        df = pd.read_excel(excel_path)
        if col_name not in df.columns or col_prov not in df.columns:
            raise KeyError(f'Excel 中未找到列：{col_name} 或 {col_prov}')
        df = df[[col_name, col_prov]].dropna()
        df[col_name] = df[col_name].astype(str).str.strip()
        df[col_prov] = df[col_prov].astype(str).str.strip()
        return dict(zip(df[col_name], df[col_prov]))

    def collect_all_files(self, root_dirs: List[str]) -> List[Path]:
        files = []
        for r in root_dirs:
            if os.path.exists(r):
                files.extend(Path(r).rglob('*'))
        return [f for f in files if f.is_file()]

    def extract_date_from_filename(self, fname: str) -> datetime:
        date_patterns = [
            r'(\d{4}-\d{2}-\d{2})',
            r'(\d{8})',
            r'(\d{4}年\d{1,2}月\d{1,2}日)'
        ]
        for pattern in date_patterns:
            m = re.search(pattern, fname)
            if m:
                date_str = m.group(1)
                try:
                    if '-' in date_str:
                        return datetime.strptime(date_str, '%Y-%m-%d')
                    elif '年' in date_str:
                        return datetime.strptime(date_str, '%Y年%m月%d日')
                    else:
                        return datetime.strptime(date_str, '%Y%m%d')
                except ValueError:
                    continue
        return datetime.fromtimestamp(os.path.getmtime(fname))

    def match_person(self, fname: str, persons: Set[str]) -> Optional[str]:
        """
        使用 fnmatch 算法匹配人名，类似 Tab1 的可靠匹配方式
        支持长文件名如 "崔晓雪_医美消费分期业务客户告知书.pdf"
        """
        base = fname.rsplit('.', 1)[0]  # 去除扩展名
        
        for p in persons:
            # 创建多种匹配模式，类似 Tab1 的方式
            patterns = [
                f"*{p}*",           # 包含人名
                f"{p}_*",           # 人名开头加下划线
                f"*_{p}",           # 下划线加人名结尾
                f"{p}*",            # 人名开头
                f"*{p}",            # 人名结尾
                f"{p} ",            # 人名开头加空格
                f" {p}",            # 空格加人名
                f"{p}-",            # 人名开头加连字符
                f"-{p}",            # 连字符加人名
            ]
            
            # 尝试所有模式
            for pattern in patterns:
                if fnmatch.fnmatch(base, pattern):
                    return p
                    
            # 也尝试完整文件名匹配（包含扩展名）
            for pattern in patterns:
                if fnmatch.fnmatch(fname, pattern):
                    return p
        
        # 调试信息：对于特定文件名，显示匹配过程
        if "双录视频" in fname:
            self.log(f"? 调试: 文件 '{fname}' 未匹配到人名，人名列表: {list(persons)}")
            for p in persons:
                for pattern in [f"*{p}*", f"{p}_*", f"*_{p}"]:
                    match_result = fnmatch.fnmatch(base, pattern)
                    self.log(f"? 测试: 人名 '{p}' 模式 '{pattern}' -> {match_result}")
                    
        return None

    def group_files_by_person(self, files: List[Path], persons: Set[str], keywords: List[str]) -> Dict[str, Dict[str, List[Tuple[Path, datetime]]]]:
        """
        返回：人名 -> 关键字 -> [(Path, date)]
        使用 fnmatch 算法进行关键字匹配，提高匹配准确性
        """
        bucket: Dict[str, Dict[str, List[Tuple[Path, datetime]]]] = {p: {} for p in persons}
        self.log(f"开始匹配文件，关键字列表: {keywords}")
        
        for f in files:
            matched_person = self.match_person(f.name, persons)
            if not matched_person:
                continue
                
            # 改进的关键字匹配逻辑
            for kw in keywords:
                # 使用多种模式匹配关键字
                kw_patterns = [
                    f"*{kw}*",           # 包含关键字
                    f"{kw}*",            # 关键字开头
                    f"*{kw}",            # 关键字结尾
                ]
                
                matched = False
                for pattern in kw_patterns:
                    if fnmatch.fnmatch(f.name, pattern):
                        matched = True
                        break
                
                if matched:
                    bucket.setdefault(matched_person, {}).setdefault(kw, [])
                    file_date = self.extract_date_from_filename(str(f))
                    bucket[matched_person][kw].append((f, file_date))
                    self.log(f"✓ 匹配成功: {f.name} -> 人名:{matched_person}, 关键字:{kw}")
                else:
                    # 调试信息：显示为什么没有匹配
                    if kw in keywords[:3]:  # 只显示前3个关键字的调试信息，避免日志过多
                        self.log(f"? 关键字不匹配: {f.name} vs {kw}")
                        
        return bucket

    def select_latest(self, grouped: Dict[str, Dict[str, List[Tuple[Path, datetime]]]]) -> Dict[str, Dict[str, Path]]:
        latest: Dict[str, Dict[str, Path]] = {}
        for person, kw_dict in grouped.items():
            latest[person] = {}
            for kw, lst in kw_dict.items():
                if lst:
                    lst.sort(key=lambda x: x[1], reverse=True)
                    latest[person][kw] = lst[0][0]
        return latest

    def copy_and_report(self, latest: Dict[str, Dict[str, Path]],
                        person_province: Dict[str, str],
                        dest_root: Path,
                        keywords: List[str],
                        file_mode: str = "copy"):
        dest_root.mkdir(parents=True, exist_ok=True)
        missing = []
        operation_count = 0
        
        for person, kw_map in latest.items():
            province = person_province[person]
            person_dir = dest_root / province / person
            person_dir.mkdir(parents=True, exist_ok=True)
            for kw in keywords:
                if kw not in kw_map:
                    missing.append(f'[MISS] {province}\\{person} 缺少：{kw}')
                    continue
                src = kw_map[kw]
                dst = person_dir / src.name
                
                # Skip if destination already exists
                if dst.exists():
                    self.log(f'[SKIP] 已存在: {dst}')
                    continue
                
                try:
                    if file_mode == "cut":
                        # 剪切模式：移动文件
                        shutil.move(str(src), str(dst))
                        self.log(f'[MOVE] {src} -> {dst}')
                    else:
                        # 复制模式：复制文件
                        shutil.copy2(src, dst)
                        self.log(f'[COPY] {src} -> {dst}')
                    operation_count += 1
                except Exception as e:
                    self.log(f'[ERROR] {file_mode.upper()} 失败: {src} -> {dst} - {e}')
        
        # 打印并写日志
        for m in missing:
            self.log(m)
        if missing:
            (dest_root / 'missing_files.log').write_text('\n'.join(missing), encoding='utf-8')
        
        operation_text = "移动" if file_mode == "cut" else "复制"
        self.log(f"{operation_text}完成: {operation_count} 个文件，缺失 {len(missing)} 个文件")

    def select_dir(self, var):
        p = filedialog.askdirectory()
        if p: var.set(os.path.normpath(p))

    def load_excel_tab3(self):
        """加载Excel文件并更新列选择器"""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            try:
                self.excel_file_var.set(path)
                self.df_tab3 = pd.read_excel(path)
                
                # 更新关联数据列下拉框
                cols = self.df_tab3.columns.tolist()
                self.association_col_cb['values'] = cols
                
                # 清空现有目录结构配置
                self.directory_structure = []
                self.update_tree_display()
                
                self.log(f"Excel 加载成功: {len(self.df_tab3)} 行, {len(cols)} 列")
            except Exception as e:
                messagebox.showerror("错误", f"Excel 读取失败: {e}")
                self.df_tab3 = None

    def start_advanced_organize_task(self):
        """开始高级整理任务"""
        excel_file = self.excel_file_var.get().strip()
        association_col = self.association_col_cb.get()
        dest_root = self.dest_root_var.get().strip()
        file_mode = self.file_mode_var.get()
        
        # 验证配置
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("错误", "Excel文件路径无效")
            return
        
        if not association_col:
            messagebox.showerror("错误", "请选择关联数据列")
            return
            
        if not self.directory_structure:
            messagebox.showerror("错误", "请至少配置一个目录层级")
            return
        
        if not dest_root:
            messagebox.showerror("错误", "请选择目标目录")
            return
        
        # 获取所有关键字
        all_keywords = self.get_all_keywords()
        if not all_keywords:
            messagebox.showerror("错误", "请至少添加一个关键字")
            return
        
        # 获取目录列配置
        directory_cols = self.get_directory_columns()
        
        self.log(f"--- 开始高级文件整理任务 ({'剪切模式' if file_mode == 'cut' else '复制模式'}) ---")
        self.log(f"关联数据列: {association_col}")
        self.log(f"目录层级: {' -> '.join(directory_cols)}")
        self.log(f"所有关键字: {all_keywords}")
        
        try:
            # 执行高级整理逻辑
            self.execute_advanced_organize(excel_file, association_col, 
                                         directory_cols, all_keywords, 
                                         Path(dest_root), file_mode)
            self.log("高级整理完成！")
            messagebox.showinfo("完成", f"高级文件整理任务完成！({'剪切' if file_mode == 'cut' else '复制'}模式)")
            
        except Exception as e:
            self.log(f"整理失败: {e}")
            messagebox.showerror("错误", f"整理失败: {e}")

    def execute_advanced_organize(self, excel_path: str, association_col: str, 
                                 directory_cols: List[str], keywords: List[str],
                                 dest_root: Path, file_mode: str):
        """执行高级整理逻辑"""
        # 1. 加载Excel数据
        df = pd.read_excel(excel_path)
        
        # 2. 收集所有源文件
        all_files = self.collect_all_files(self.root_dirs_list)
        self.log(f"共发现 {len(all_files)} 个原始文件")
        
        # 3. 文件匹配和分组
        processed_count = 0
        for _, row in df.iterrows():
            # 获取关联数据值（用于文件名匹配）
            association_value = str(row[association_col]).strip()
            if not association_value or association_value == 'nan':
                continue
            
            # 构建目录路径
            dir_parts = []
            for col in directory_cols:
                value = str(row[col]).strip()
                if value and value != 'nan':
                    # 清理文件名非法字符
                    clean_value = value
                    for char in r'[\/:*?"<>|]': 
                        clean_value = clean_value.replace(char, "_")
                    dir_parts.append(clean_value)
            
            if not dir_parts:
                continue
            
            # 创建目标目录
            target_dir = dest_root
            for part in dir_parts:
                target_dir = target_dir / part
            target_dir.mkdir(parents=True, exist_ok=True)
            
            # 4. 匹配并复制文件
            matched_files = self.match_files_for_association(all_files, association_value, keywords)
            
            for keyword, file_path in matched_files.items():
                if file_path:
                    dest_file = target_dir / file_path.name
                    
                    # 跳过已存在的文件
                    if dest_file.exists():
                        self.log(f"[SKIP] 已存在: {dest_file}")
                        continue
                    
                    # 执行文件操作
                    try:
                        if file_mode == "cut":
                            shutil.move(str(file_path), str(dest_file))
                            self.log(f"[MOVE] {file_path.name} -> {target_dir}")
                        else:
                            shutil.copy2(str(file_path), str(dest_file))
                            self.log(f"[COPY] {file_path.name} -> {target_dir}")
                        processed_count += 1
                    except Exception as e:
                        self.log(f"[ERROR] 处理文件失败: {file_path.name} - {e}")
        
        self.log(f"处理完成，共处理 {processed_count} 个文件")

    def match_files_for_association(self, files: List[Path], association_value: str, keywords: List[str]) -> Dict[str, Optional[Path]]:
        """为关联值匹配文件，返回 {关键字: 文件路径}"""
        matched_files = {}
        
        for file_path in files:
            # 检查文件名是否包含关联值
            if not self.match_person(file_path.name, {association_value}):
                continue
            
            # 检查关键字匹配
            for keyword in keywords:
                if keyword in matched_files:
                    continue  # 已找到该关键字的文件
                
                # 使用改进的关键字匹配
                kw_patterns = [f"*{keyword}*", f"{keyword}*", f"*{keyword}"]
                for pattern in kw_patterns:
                    if fnmatch.fnmatch(file_path.name, pattern):
                        matched_files[keyword] = file_path
                        break
        
        return matched_files

    # ================= Tab 4: 智能文件归档 =================
    def setup_tab4(self):
        """Tab4: 智能文件归档 - 基于Excel数据驱动的文件自动归档"""
        # Tab4变量初始化
        self.tab4_excel_path_var = tk.StringVar()
        self.tab4_key_col_var = tk.StringVar()
        self.tab4_key_col_optional_var = tk.StringVar()
        # 文件夹命名列（3个）
        self.tab4_folder_name_col1_var = tk.StringVar()
        self.tab4_folder_name_col2_var = tk.StringVar()
        self.tab4_folder_name_col3_var = tk.StringVar()
        # 文件名命名列（3个）
        self.tab4_file_name_col1_var = tk.StringVar()
        self.tab4_file_name_col2_var = tk.StringVar()
        self.tab4_file_name_col3_var = tk.StringVar()
        self.tab4_source_dir_var = tk.StringVar()
        self.tab4_dest_dir_var = tk.StringVar()
        self.tab4_ocr_pages_var = tk.StringVar(value="1")
        self.tab4_title_page_var = tk.StringVar(value="1")  # 标题所在页，默认第1页
        # OCR输出限制和提示词配置
        self.tab4_ocr_max_chars_var = tk.StringVar(value="")  # 识别文字量（可选）
        self.tab4_ocr_prompt_var = tk.StringVar(value="")  # 补充提示词（可选）
        # Tab4独立的火山引擎配置（不使用Tab2的）
        self.tab4_api_key_var = tk.StringVar(value=self.config.get("tab4_api_key", ""))
        self.tab4_model_id_var = tk.StringVar(value=self.config.get("tab4_model_id", "doubao-seed-1-6-vision-250815"))
        self.tab4_df = None
        self.tab4_processed_files = set()
        self.tab4_is_running = False
        self.tab4_stop_flag = False
        
        frame = self.tab4
        frame.configure(bg=COLORS['bg_primary'])
        
        # 创建滚动容器
        main_canvas = tk.Canvas(frame, bg=COLORS['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=COLORS['bg_primary'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        self.scrollable_window_tab4 = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        def on_canvas_configure(event):
            canvas_width = event.width - 4
            main_canvas.itemconfig(self.scrollable_window_tab4, width=canvas_width)
        main_canvas.bind('<Configure>', on_canvas_configure)
        
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            main_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            main_canvas.unbind_all("<MouseWheel>")
        
        main_canvas.bind('<Enter>', bind_mousewheel)
        main_canvas.bind('<Leave>', unbind_mousewheel)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 标题
        title_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="智能文件归档",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="基于Excel数据，自动识别文件内容并归档到指定目录",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))
        
        # 1. Excel数据源卡片
        excel_card = self.create_card(scrollable_frame, "Excel数据源")
        excel_card.pack(fill="x", pady=8)
        
        # Excel文件选择
        excel_row = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        excel_row.pack(fill="x", pady=(0, 12))
        tk.Entry(excel_row, 
                textvariable=self.tab4_excel_path_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(excel_row, "加载Excel", self.load_excel_tab4).pack(side="right")
        
        # 列选择区域
        col_grid = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        col_grid.pack(fill="x")
        
        # 关键数据列（必填）
        tk.Label(col_grid, 
                text="关键数据列*:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        self.tab4_key_col_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_key_col_cb.grid(row=0, column=1, padx=(8, 16), sticky="w")
        
        # 关键数据列（可选）
        tk.Label(col_grid, 
                text="关键数据列(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="e", pady=6)
        self.tab4_key_col_opt_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_key_col_opt_cb.grid(row=0, column=3, padx=(8, 0), sticky="w")
        
        # 文件夹命名列组
        folder_label = tk.Label(col_grid, 
                text="【文件夹命名规则】",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold'))
        folder_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(12, 4))
        
        # 文件夹命名列1（必填）
        tk.Label(col_grid, 
                text="文件夹列1*:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=0, sticky="e", pady=6)
        self.tab4_folder_name_col1_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_folder_name_col1_cb.grid(row=2, column=1, padx=(8, 16), sticky="w")
        
        # 文件夹命名列2（可选，含文件标题选项）
        tk.Label(col_grid, 
                text="文件夹列2(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=2, sticky="e", pady=6)
        self.tab4_folder_name_col2_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_folder_name_col2_cb.grid(row=2, column=3, padx=(8, 0), sticky="w")
        
        # 文件夹命名列3（可选，含文件标题选项）
        tk.Label(col_grid, 
                text="文件夹列3(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=3, column=2, sticky="e", pady=6)
        self.tab4_folder_name_col3_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_folder_name_col3_cb.grid(row=3, column=3, padx=(8, 0), sticky="w")
        
        # 文件名命名列组
        file_label = tk.Label(col_grid, 
                text="【复制后文件名命名规则】",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold'))
        file_label.grid(row=4, column=0, columnspan=4, sticky="w", pady=(12, 4))
        
        # 文件名命名列1
        tk.Label(col_grid, 
                text="文件名列1:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=5, column=0, sticky="e", pady=6)
        self.tab4_file_name_col1_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_file_name_col1_cb.grid(row=5, column=1, padx=(8, 16), sticky="w")
        
        # 文件名命名列2（可选，含文件标题和原文件名选项）
        tk.Label(col_grid, 
                text="文件名列2(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=5, column=2, sticky="e", pady=6)
        self.tab4_file_name_col2_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_file_name_col2_cb.grid(row=5, column=3, padx=(8, 0), sticky="w")
        
        # 文件名命名列3（可选，含文件标题和原文件名选项）
        tk.Label(col_grid, 
                text="文件名列3(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=6, column=2, sticky="e", pady=6)
        self.tab4_file_name_col3_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab4_file_name_col3_cb.grid(row=6, column=3, padx=(8, 0), sticky="w")
        
        # 2. OCR配置卡片
        ocr_card = self.create_card(scrollable_frame, "OCR配置")
        ocr_card.pack(fill="x", pady=8)
        
        ocr_grid = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        ocr_grid.pack(fill="x")
        
        # OCR页码配置 - 第一行：关键数据识别页码
        tk.Label(ocr_grid, 
                text="关键数据识别页码:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab4_ocr_pages_var,
                width=15,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(支持多页码，如: 1,2,3 或 1-3,5)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="w")
        
        # 标题所在页配置 - 第二行
        tk.Label(ocr_grid, 
                text="标题所在页:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab4_title_page_var,
                width=5,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=1, column=1, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(第几页识别文件标题，用于命名列中的[文件标题])",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=2, sticky="w")
        
        # 识别文字量配置 - 第三行
        tk.Label(ocr_grid, 
                text="识别文字量(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab4_ocr_max_chars_var,
                width=10,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=2, column=1, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(LLM OCR时最大输出字数，如: 100)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=2, sticky="w")
        
        # 补充提示词配置 - 第四行
        tk.Label(ocr_grid, 
                text="补充提示词(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=3, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab4_ocr_prompt_var,
                width=40,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=3, column=1, columnspan=2, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(传给LLM的额外提示，精简返回关键信息)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=4, column=2, sticky="w")
        
        # 火山引擎AI配置 - 第三行开始
        ai_config_frame = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        ai_config_frame.pack(fill="x", pady=(12, 0))
        
        tk.Label(ai_config_frame, 
                text="火山引擎AI配置 (LLM OCR兜底)",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold')).pack(anchor="w", pady=(0, 8))
        
        ai_grid = tk.Frame(ai_config_frame, bg=COLORS['bg_primary'])
        ai_grid.pack(fill="x")
        
        # API Key
        tk.Label(ai_grid, 
                text="API Key:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(ai_grid, 
                textvariable=self.tab4_api_key_var, 
                show="●",
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        self.create_secondary_button(ai_grid, "测试连接", self.test_ai_connection_tab4).grid(row=0, column=2)
        
        # Model ID
        tk.Label(ai_grid, 
                text="模型ID:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Entry(ai_grid, 
                textvariable=self.tab4_model_id_var, 
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=1, column=1, padx=(12, 8), sticky="w")
        
        # OCR模式说明
        mode_frame = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        mode_frame.pack(fill="x", pady=(12, 0))
        tk.Label(mode_frame, 
                text="OCR模式: 本地OCR (RapidOCR) → LLM兜底 (火山引擎，需配置API Key)",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9)).pack(anchor="w")
        
        # 3. 文件夹配置卡片
        folder_card = self.create_card(scrollable_frame, "文件夹配置")
        folder_card.pack(fill="x", pady=8)
        
        # 文件源
        src_row = tk.Frame(folder_card, bg=COLORS['bg_primary'])
        src_row.pack(fill="x", pady=(0, 8))
        tk.Label(src_row, 
                text="文件源:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(src_row, 
                textvariable=self.tab4_source_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8), fill="x", expand=True)
        self.create_secondary_button(src_row, "浏览...", lambda: self.select_dir(self.tab4_source_dir_var)).pack(side="right")
        
        # 归档目的地
        dest_row = tk.Frame(folder_card, bg=COLORS['bg_primary'])
        dest_row.pack(fill="x")
        tk.Label(dest_row, 
                text="归档目的地:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(dest_row, 
                textvariable=self.tab4_dest_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8), fill="x", expand=True)
        self.create_secondary_button(dest_row, "浏览...", lambda: self.select_dir(self.tab4_dest_dir_var)).pack(side="right")
        
        # 4. 执行按钮区域
        action_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'], pady=20)
        action_frame.pack(fill="x", side="bottom")
        
        button_container = tk.Frame(action_frame, bg=COLORS['bg_primary'])
        button_container.pack(fill="x")
        
        self.tab4_start_btn = self.create_primary_button(button_container, "▶ 开始智能归档", self.start_archive_task)
        self.tab4_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.tab4_stop_btn = tk.Button(button_container, 
                                     text="■ 停止任务", 
                                     command=self.stop_archive_task,
                                     bg=COLORS['danger'],
                                     fg='white',
                                     font=('微软雅黑', 10, 'bold'),
                                     relief='flat',
                                     padx=20,
                                     pady=8,
                                     cursor='hand2',
                                     state="disabled")
        self.tab4_stop_btn.pack(side="right", fill="x", expand=True)

    def load_excel_tab4(self):
        """加载Excel文件并更新列选择器"""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            try:
                self.tab4_excel_path_var.set(path)
                self.tab4_df = pd.read_excel(path)
                cols = self.tab4_df.columns.tolist()
                cols_with_empty = [""] + cols
                cols_with_title = ["", "[文件标题]"] + cols
                # 文件名命名列包含原文件名选项
                cols_with_title_and_original = ["", "[文件标题]", "[原文件名]"] + cols
                
                self.tab4_key_col_cb['values'] = cols
                self.tab4_key_col_opt_cb['values'] = cols_with_empty
                
                # 文件夹命名列（不含原文件名选项）
                self.tab4_folder_name_col1_cb['values'] = cols
                self.tab4_folder_name_col2_cb['values'] = cols_with_title
                self.tab4_folder_name_col3_cb['values'] = cols_with_title
                
                # 文件名命名列（含原文件名选项）
                self.tab4_file_name_col1_cb['values'] = cols_with_title_and_original
                self.tab4_file_name_col2_cb['values'] = cols_with_title_and_original
                self.tab4_file_name_col3_cb['values'] = cols_with_title_and_original
                
                self.log(f"Tab4 Excel加载成功: {len(self.tab4_df)} 行, {len(cols)} 列")
            except Exception as e:
                messagebox.showerror("错误", f"Excel读取失败: {e}")
                self.tab4_df = None

    def test_ai_connection_tab4(self):
        """测试Tab4火山引擎连接"""
        key = self.tab4_api_key_var.get().strip()
        m_id = self.tab4_model_id_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请填写 API KEY")
            return
        
        # 保存配置
        self.config["tab4_api_key"] = key
        self.config["tab4_model_id"] = m_id
        self.save_config()
        
        self.log(f"正在测试Tab4连接: {m_id} ...")
        
        try:
            client = OpenAI(api_key=key, base_url="https://ark.cn-beijing.volces.com/api/v3")
            response = client.chat.completions.create(
                model=m_id,
                messages=[{"role": "user", "content": "hi"}],
                max_tokens=5
            )
            self.log("√ Tab4连接成功！模型响应正常。")
            messagebox.showinfo("成功", "连接成功！")
        except Exception as e:
            self.log(f"× Tab4连接失败: {e}")
            messagebox.showerror("连接失败", f"错误信息:\n{e}")



    def test_ai_connection_tab5(self):
        """测试Tab5独立OCR模型连接"""
        key = self.tab5_ocr_api_key_var.get().strip()
        m_id = self.tab5_ocr_model_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请填写 API KEY")
            return
            
        self.config["tab5_ocr_api_key"] = key
        self.config["tab5_ocr_model_id"] = m_id
        self.save_config()
        
        self.log(f"正在测试Tab5 OCR连接: {m_id} ...")
        
        try:
            from openai import OpenAI
            client = OpenAI(
                api_key=key,
                base_url="https://ark.cn-beijing.volces.com/api/v3"
            )
            response = client.chat.completions.create(
                model=m_id,
                messages=[{"role": "user", "content": "hi"}],
                max_tokens=5
            )
            
            messagebox.showinfo("成功", f"连接成功！\n模型响应: {response.choices[0].message.content}")
            self.log(f"✓ Tab5 OCR连接成功: {response.choices[0].message.content}")
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("连接失败", f"错误信息:\n{error_msg}")
            self.log(f"× Tab5 OCR连接失败: {error_msg}")

    def test_ai_connection_tab5(self):
        """测试Tab5独立OCR模型连接"""
        key = self.tab5_ocr_api_key_var.get().strip()
        m_id = self.tab5_ocr_model_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请填写 API KEY")
            return
            
        self.config["tab5_ocr_api_key"] = key
        self.config["tab5_ocr_model_id"] = m_id
        self.save_config()
        
        self.log(f"正在测试Tab5 OCR连接: {m_id} ...")
        
        try:
            from openai import OpenAI
            client = OpenAI(
                api_key=key,
                base_url="https://ark.cn-beijing.volces.com/api/v3"
            )
            response = client.chat.completions.create(
                model=m_id,
                messages=[{"role": "user", "content": "hi"}],
                max_tokens=5
            )
            
            # Record tokens
            if hasattr(response, 'usage') and response.usage:
                prompt_tokens = getattr(response.usage, 'prompt_tokens', 0)
                completion_tokens = getattr(response.usage, 'completion_tokens', 0)
                total_tokens = getattr(response.usage, 'total_tokens', 0)
                self.log(f"Token消耗 - 输入:{prompt_tokens}, 输出:{completion_tokens}, 总计:{total_tokens}")
            
            messagebox.showinfo("成功", f"连接成功！\n模型响应: {response.choices[0].message.content}")
            self.log(f"✓ Tab5 OCR连接成功: {response.choices[0].message.content}")
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("连接失败", f"错误信息:\n{error_msg}")
            self.log(f"× Tab5 OCR连接失败: {error_msg}")

    def parse_page_numbers(self, page_str):
        """解析页码字符串，如 '1,2,3' 或 '1-3,5' -> [1,2,3,5]"""
        pages = set()
        try:
            parts = page_str.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start, end = part.split('-', 1)
                    start, end = int(start.strip()), int(end.strip())
                    pages.update(range(start, end + 1))
                else:
                    pages.add(int(part.strip()))
            return sorted(list(pages))
        except:
            return [1]

    def stop_archive_task(self):
        """停止归档任务"""
        if self.tab4_is_running:
            self.tab4_stop_flag = True
            self.log("🛑 正在停止归档任务...")
            self.tab4_stop_btn.configure(state="disabled", text="停止中...")

    def update_tab4_ui_state(self, is_running):
        """更新Tab4 UI状态"""
        self.tab4_is_running = is_running
        if is_running:
            self.tab4_start_btn.configure(state="disabled")
            self.tab4_stop_btn.configure(state="normal", text="■ 停止任务")
            self.tab4_stop_flag = False
        else:
            self.tab4_start_btn.configure(state="normal")
            self.tab4_stop_btn.configure(state="disabled", text="■ 停止任务")

    def start_archive_task(self):
        """启动智能归档任务"""
        if self.tab4_df is None:
            messagebox.showwarning("配置不全", "请先加载Excel文件")
            return

        key_col = self.tab4_key_col_cb.get()
        folder_col1 = self.tab4_folder_name_col1_cb.get()

        if not key_col:
            messagebox.showwarning("配置不全", "请选择关键数据列")
            return
        if not folder_col1:
            messagebox.showwarning("配置不全", "请至少选择文件夹命名列1")
            return

        source_dir = self.tab4_source_dir_var.get().strip()
        dest_dir = self.tab4_dest_dir_var.get().strip()

        if not source_dir or not os.path.exists(source_dir):
            messagebox.showerror("错误", "文件源路径无效")
            return
        if not dest_dir:
            messagebox.showerror("错误", "请选择归档目的地")
            return

        page_nums = self.parse_page_numbers_tab4(self.tab4_ocr_pages_var.get())

        try:
            title_page = int(self.tab4_title_page_var.get().strip())
        except ValueError:
            title_page = 1

        folder_cols = [
            folder_col1,
            self.tab4_folder_name_col2_cb.get(),
            self.tab4_folder_name_col3_cb.get()
        ]
        file_cols = [
            self.tab4_file_name_col1_cb.get(),
            self.tab4_file_name_col2_cb.get(),
            self.tab4_file_name_col3_cb.get()
        ]

        ocr_max_chars = self.tab4_ocr_max_chars_var.get().strip()
        ocr_prompt = self.tab4_ocr_prompt_var.get().strip()

        thread = threading.Thread(
            target=self._run_archive_task,
            args=(key_col, folder_cols, file_cols, page_nums, title_page, ocr_max_chars, ocr_prompt),
            daemon=True
        )
        thread.start()

    def parse_page_numbers_tab4(self, page_str):
        """解析Tab4页码字符串"""
        pages = set()
        try:
            parts = page_str.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start, end = part.split('-', 1)
                    start, end = int(start.strip()), int(end.strip())
                    pages.update(range(start, end + 1))
                else:
                    pages.add(int(part.strip()))
            return sorted(list(pages))
        except:
            return [1]

    def _sanitize_folder_name_tab4(self, name):
        """清理Tab4文件夹名称中的非法字符"""
        if not name or name == 'nan':
            return "unnamed"
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            name = name.replace(char, '_')
        name = ''.join(char for char in name if ord(char) >= 32)
        name = name.strip(' .')
        if not name:
            name = "unnamed"
        return name

    def _get_file_content_with_pages_tab4(self, file_path, page_nums):
        """获取Tab4文件多页内容（PDF或图片）"""
        contents = []
        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.pdf':
                doc = fitz.open(file_path)
                total_pages = len(doc)

                for page_num in page_nums:
                    if page_num < 1 or page_num > total_pages:
                        continue
                    page = doc.load_page(page_num - 1)
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")
                    contents.append((page_num, img_bytes))

                doc.close()
            elif ext in ['.jpg', '.jpeg', '.png']:
                with open(file_path, 'rb') as f:
                    img_bytes = f.read()
                contents.append((1, img_bytes))
        except Exception as e:
            self.log(f"  文件读取错误: {e}")

        return contents

    def _ocr_local_tab4(self, img_bytes, file_path):
        """Tab4本地OCR识别"""
        try:
            if self.tab4_ocr_engine is None:
                self.log("  正在初始化Tab4本地OCR引擎(RapidOCR)，请稍候...")
                try:
                    ocr_result = [None]
                    def init_ocr():
                        try:
                            ocr_result[0] = RapidOCR()
                        except Exception as e:
                            self.log(f"  Tab4 OCR引擎初始化失败: {e}")
                            ocr_result[0] = None

                    ocr_thread = threading.Thread(target=init_ocr)
                    ocr_thread.daemon = True
                    ocr_thread.start()
                    ocr_thread.join(timeout=30)

                    if ocr_result[0] is None:
                        self.log("  ⚠️ Tab4 OCR引擎初始化超时或失败，跳过本地OCR")
                        return ""

                    self.tab4_ocr_engine = ocr_result[0]
                    self.log("  ✓ Tab4本地OCR引擎初始化完成")
                except Exception as e:
                    self.log(f"  ⚠️ Tab4 OCR初始化异常: {e}")
                    return ""

            result, _ = self.tab4_ocr_engine(img_bytes)
            if result:
                text = "".join([line[1] for line in result])
                self.log(f"  ✓ OCR识别完成，提取 {len(text)} 字符")
                return text
            else:
                self.log("  ⚠️ OCR未识别到文字")
            return ""
        except Exception as e:
            self.log(f"  ⚠️ Tab4本地OCR错误: {e}")
            return ""

    def _ocr_llm_fallback_tab4(self, img_bytes, ocr_max_chars="", ocr_prompt=""):
        """Tab4 LLM兜底OCR"""
        api_key = self.tab4_api_key_var.get().strip()
        if not api_key:
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

        try:
            client = OpenAI(api_key=api_key, base_url="https://ark.cn-beijing.volces.com/api/v3")
            base64_image = base64.b64encode(img_bytes).decode('utf-8')

            if ocr_prompt:
                prompt = ocr_prompt
            else:
                prompt = "请提取图片中的文字内容。直接输出识别到的文字，不要包含解释或格式标记。"

            if ocr_max_chars:
                try:
                    max_chars = int(ocr_max_chars)
                    prompt += f"\n请控制在{max_chars}个字符以内。"
                except:
                    pass

            request_params = {
                "model": self.tab4_model_id_var.get(),
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                            }
                        ]
                    }
                ]
            }

            if ocr_max_chars:
                try:
                    request_params["max_tokens"] = int(int(ocr_max_chars) * 2)
                except:
                    pass

            response = client.chat.completions.create(**request_params)
            result_text = response.choices[0].message.content

            token_usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
            if hasattr(response, 'usage') and response.usage:
                token_usage = {
                    "prompt_tokens": getattr(response.usage, 'prompt_tokens', 0),
                    "completion_tokens": getattr(response.usage, 'completion_tokens', 0),
                    "total_tokens": getattr(response.usage, 'total_tokens', 0)
                }
            return result_text, token_usage
        except Exception as e:
            self.log(f"  Tab4 LLM OCR错误: {e}")
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

    def _extract_file_title(self, file_path, title_page=1):
        """提取文件标题（PDF使用Tab2逻辑，图片返回空）"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.pdf':
            return self.extract_pdf_title(file_path, title_page)
        return ""

    def _build_folder_name(self, row_data, name_cols, file_title=""):
        """构建文件夹名称"""
        parts = []
        for col in name_cols:
            if not col:
                continue
            if col == "[文件标题]":
                if file_title:
                    parts.append(self._sanitize_folder_name_tab4(file_title))
            else:
                value = str(row_data.get(col, "")).strip()
                if value and value != 'nan':
                    parts.append(self._sanitize_folder_name_tab4(value))
        
        return "_".join(parts) if parts else "unnamed"

    def _match_content_with_keywords(self, content, keywords):
        """检查内容是否包含任一关键字"""
        content_norm = self._normalize_for_matching(content)
        for kw in keywords:
            kw_norm = self._normalize_for_matching(kw)
            if kw_norm and kw_norm in content_norm:
                return kw
        return None

    def _run_archive_task(self, key_col, folder_cols, file_cols, page_nums, title_page, ocr_max_chars="", ocr_prompt=""):
        """执行智能归档任务"""
        try:
            self.update_tab4_ui_state(True)
            
            source_dir = self.tab4_source_dir_var.get()
            dest_dir = self.tab4_dest_dir_var.get()
            key_col_opt = self.tab4_key_col_opt_cb.get()
            
            self.log(f"--- 开始智能归档任务 ---")
            self.log(f"关键数据列: {key_col}")
            self.log(f"文件夹命名列: {folder_cols}")
            self.log(f"文件名命名列: {file_cols}")
            self.log(f"OCR页码: {page_nums}")
            self.log(f"标题识别页: 第{title_page}页")
            if ocr_max_chars:
                self.log(f"LLM文字量限制: {ocr_max_chars}字")
            if ocr_prompt:
                self.log(f"LLM补充提示词: {ocr_prompt}")
            
            # 1. 构建数据映射 {关键值: 行数据}
            data_map = {}
            keywords = []
            for _, row in self.tab4_df.iterrows():
                if self.tab4_stop_flag:
                    break
                key_val = str(row.get(key_col, "")).strip()
                if key_val and key_val != 'nan':
                    keywords.append(key_val)
                    data_map[key_val] = row
            
            self.log(f"从Excel提取 {len(keywords)} 个关键数据")
            
            # 辅助函数：构建文件名
            def build_file_name(row_data, file_cols, original_name, file_title=""):
                parts = []
                for col in file_cols:
                    if not col:
                        continue
                    if col == "[文件标题]":
                        if file_title:
                            parts.append(self._sanitize_folder_name_tab4(file_title))
                    elif col == "[原文件名]":
                        name_without_ext = os.path.splitext(original_name)[0]
                        parts.append(name_without_ext)
                    else:
                        value = str(row_data.get(col, "")).strip()
                        if value and value != 'nan':
                            parts.append(self._sanitize_folder_name_tab4(value))
                if parts:
                    ext = os.path.splitext(original_name)[1]
                    return "_".join(parts) + ext
                return original_name
            
            # 辅助函数：移除OCR文本中的空格（处理身份证号等连续数据）
            def normalize_ocr_text(text):
                """移除OCR识别结果中的空格，保留原始文本供记录"""
                if not text:
                    return ""
                # 移除所有空格，用于匹配
                return text.replace(" ", "").replace("\t", "")
            
            # 2. 收集所有PDF/JPG/PNG文件
            valid_exts = {'.pdf', '.jpg', '.jpeg', '.png'}
            all_files = []
            for root, dirs, files in os.walk(source_dir):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in valid_exts:
                        all_files.append(os.path.join(root, f))
            
            self.log(f"发现 {len(all_files)} 个待处理文件")
            
            # 3. 第一轮：文件名匹配
            processed_files = set()
            unmatched_files = []
            
            for file_path in all_files:
                if self.tab4_stop_flag:
                    break
                
                file_name = os.path.basename(file_path)
                matched = False
                
                for keyword in keywords:
                    # 文件名匹配（支持多种模式）
                    patterns = [f"*{keyword}*", f"{keyword}_*", f"*_{keyword}", f"{keyword}*", f"*{keyword}"]
                    for pattern in patterns:
                        if fnmatch.fnmatch(file_name, pattern) or fnmatch.fnmatch(os.path.splitext(file_name)[0], pattern):
                            # 匹配成功，执行归档
                            row = data_map[keyword]
                            file_title = self._extract_file_title(file_path, title_page) if "[文件标题]" in folder_cols or "[文件标题]" in file_cols else ""
                            folder_name = self._build_folder_name(row, folder_cols, file_title)
                            new_file_name = build_file_name(row, file_cols, file_name, file_title)
                            
                            target_folder = os.path.join(dest_dir, folder_name)
                            os.makedirs(target_folder, exist_ok=True)
                            
                            dest_file = os.path.join(target_folder, new_file_name)
                            if not os.path.exists(dest_file):
                                shutil.copy2(file_path, dest_file)
                                self.log(f"✓ 文件名匹配归档: {file_name} -> {folder_name}/{new_file_name}")
                            else:
                                self.log(f"⚠ 文件已存在: {new_file_name}")
                            
                            processed_files.add(file_path)
                            matched = True
                            break
                    if matched:
                        break
                
                if not matched:
                    unmatched_files.append(file_path)
            
            self.log(f"文件名匹配完成: {len(processed_files)} 个文件已归档, {len(unmatched_files)} 个待处理")
            
            # 4. 将未匹配文件复制到OCR待处理文件夹（保持原目录结构）
            ocr_pending_dir = os.path.join(dest_dir, "OCR待处理")
            ocr_pending_files = []
            # 存储每个文件的OCR内容，用于最后生成Excel
            file_ocr_content_map = {}  # {file_path: ocr_text}
            
            if unmatched_files and not self.tab4_stop_flag:
                self.log(f"将 {len(unmatched_files)} 个未匹配文件复制到'OCR待处理'文件夹...")
                
                for file_path in unmatched_files:
                    if self.tab4_stop_flag:
                        break
                    
                    # 计算相对路径保持目录结构
                    rel_path = os.path.relpath(file_path, source_dir)
                    dest_path = os.path.join(ocr_pending_dir, rel_path)
                    
                    # 创建目标目录
                    dest_folder = os.path.dirname(dest_path)
                    os.makedirs(dest_folder, exist_ok=True)
                    
                    file_name = os.path.basename(file_path)
                    
                    # 处理文件名冲突
                    counter = 1
                    final_dest = dest_path
                    while os.path.exists(final_dest):
                        name, ext = os.path.splitext(dest_path)
                        final_dest = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    shutil.copy2(file_path, final_dest)
                    ocr_pending_files.append(final_dest)
                    self.log(f"  复制到OCR待处理: {file_name}")
                
                self.log(f"✓ 已复制 {len(ocr_pending_files)} 个文件到OCR待处理文件夹")
            
            # 5. 第二轮：批量本地OCR识别
            local_ocr_matched = []
            local_ocr_unmatched = []
            
            if ocr_pending_files and not self.tab4_stop_flag:
                self.log(f"开始批量本地OCR识别 {len(ocr_pending_files)} 个文件...")
                
                for file_path in ocr_pending_files:
                    if self.tab4_stop_flag:
                        break
                    
                    file_name = os.path.basename(file_path)
                    contents = self._get_file_content_with_pages_tab4(file_path, page_nums)
                    
                    matched = False
                    best_ocr_text = ""  # 记录最佳OCR结果
                    
                    for page_num, img_bytes in contents:
                        if self.tab4_stop_flag:
                            break
                        
                        # 本地OCR
                        text = self._ocr_local_tab4(img_bytes, file_path)
                        if text:
                            best_ocr_text = text
                            # 使用无空格版本进行匹配
                            text_no_spaces = normalize_ocr_text(text)
                            matched_kw = self._match_content_with_keywords(text_no_spaces, keywords)
                            
                            if matched_kw:
                                row = data_map[matched_kw]
                                file_title = self._extract_file_title(file_path, title_page) if "[文件标题]" in folder_cols or "[文件标题]" in file_cols else ""
                                folder_name = self._build_folder_name(row, folder_cols, file_title)
                                new_file_name = build_file_name(row, file_cols, file_name, file_title)
                                
                                target_folder = os.path.join(dest_dir, folder_name)
                                os.makedirs(target_folder, exist_ok=True)
                                
                                dest_file = os.path.join(target_folder, new_file_name)
                                if not os.path.exists(dest_file):
                                    shutil.move(file_path, dest_file)
                                    self.log(f"✓ 本地OCR匹配归档 (第{page_num}页): {file_name} -> {folder_name}/{new_file_name}")
                                else:
                                    os.remove(file_path)
                                    self.log(f"⚠ 目标文件已存在，删除源文件: {file_name}")
                                
                                processed_files.add(file_path)
                                matched = True
                                break
                    
                    if matched:
                        local_ocr_matched.append(file_path)
                    else:
                        local_ocr_unmatched.append((file_path, best_ocr_text))
                        file_ocr_content_map[file_path] = best_ocr_text
                
                self.log(f"本地OCR完成: {len(local_ocr_matched)} 个匹配, {len(local_ocr_unmatched)} 个待LLM处理")
                
                # 生成本地OCR结果Excel文件
                try:
                    local_ocr_records = []
                    # 添加匹配的文件
                    for file_path in local_ocr_matched:
                        file_name = os.path.basename(file_path)
                        # 从map中获取OCR内容（如果存在）
                        ocr_text = file_ocr_content_map.get(file_path, "")
                        local_ocr_records.append({
                            "文件名": file_name,
                            "本地OCR内容": ocr_text if ocr_text else "(匹配成功但无OCR记录)",
                            "匹配状态": "匹配成功"
                        })
                    # 添加未匹配的文件
                    for file_path, ocr_text in local_ocr_unmatched:
                        file_name = os.path.basename(file_path)
                        local_ocr_records.append({
                            "文件名": file_name,
                            "本地OCR内容": ocr_text if ocr_text else "(无识别结果)",
                            "匹配状态": "待LLM处理"
                        })
                    
                    if local_ocr_records:
                        # 创建0-OCR识别结果文件夹
                        ocr_result_dir = os.path.join(dest_dir, "0-OCR识别结果")
                        os.makedirs(ocr_result_dir, exist_ok=True)
                        local_ocr_excel_path = os.path.join(ocr_result_dir, "本地OCR识别结果.xlsx")
                        df_local_ocr = pd.DataFrame(local_ocr_records)
                        df_local_ocr.to_excel(local_ocr_excel_path, index=False)
                        self.log(f"✓ 已生成本地OCR识别结果Excel: {local_ocr_excel_path}")
                except Exception as e:
                    self.log(f"⚠ 生成本地OCR Excel记录失败: {e}")
            
            # 6. 第三轮：对本地OCR未匹配的文件批量使用LLM兜底
            llm_ocr_matched = []
            llm_ocr_unmatched = []
            # 初始化Token使用统计
            token_usage_total_tab4 = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

            if local_ocr_unmatched and not self.tab4_stop_flag:
                self.log(f"开始批量LLM OCR识别 {len(local_ocr_unmatched)} 个文件...")
                
                for file_path, _ in local_ocr_unmatched:
                    if self.tab4_stop_flag:
                        break
                    
                    file_name = os.path.basename(file_path)
                    contents = self._get_file_content_with_pages_tab4(file_path, page_nums)
                    
                    matched = False
                    best_llm_text = ""
                    
                    for page_num, img_bytes in contents:
                        if self.tab4_stop_flag:
                            break
                        
                        # LLM OCR兜底
                        text, token_usage = self._ocr_llm_fallback_tab4(img_bytes, ocr_max_chars=ocr_max_chars, ocr_prompt=ocr_prompt)
                        # 累加token使用
                        for key in token_usage_total_tab4:
                            token_usage_total_tab4[key] += token_usage.get(key, 0)
                        for key in token_usage_total:
                            token_usage_total[key] += token_usage.get(key, 0)
                        if token_usage['total_tokens'] > 0:
                            self.log(f"📊 Token消耗: {token_usage['prompt_tokens']}输入 / {token_usage['completion_tokens']}输出 / {token_usage['total_tokens']}总计")
                        
                        if text:
                            best_llm_text = text
                            # 使用无空格版本进行匹配
                            text_no_spaces = normalize_ocr_text(text)
                            matched_kw = self._match_content_with_keywords(text_no_spaces, keywords)
                            
                            if matched_kw:
                                row = data_map[matched_kw]
                                file_title = self._extract_file_title(file_path, title_page) if "[文件标题]" in folder_cols or "[文件标题]" in file_cols else ""
                                folder_name = self._build_folder_name(row, folder_cols, file_title)
                                new_file_name = build_file_name(row, file_cols, file_name, file_title)
                                
                                target_folder = os.path.join(dest_dir, folder_name)
                                os.makedirs(target_folder, exist_ok=True)
                                
                                dest_file = os.path.join(target_folder, new_file_name)
                                if not os.path.exists(dest_file):
                                    shutil.move(file_path, dest_file)
                                    self.log(f"✓ LLM OCR匹配归档 (第{page_num}页): {file_name} -> {folder_name}/{new_file_name}")
                                else:
                                    os.remove(file_path)
                                    self.log(f"⚠ 目标文件已存在，删除源文件: {file_name}")
                                
                                processed_files.add(file_path)
                                matched = True
                                break
                    
                    if matched:
                        llm_ocr_matched.append(file_path)
                    else:
                        llm_ocr_unmatched.append((file_path, best_llm_text))
                        # 更新OCR内容为LLM结果（如果LLM有返回）
                        if best_llm_text:
                            file_ocr_content_map[file_path] = best_llm_text
                
                self.log(f"LLM OCR完成: {len(llm_ocr_matched)} 个匹配, {len(llm_ocr_unmatched)} 个未匹配")
            
            # 7. 第四轮：OCR仍未匹配的文件移动到0未处理文件夹
            if llm_ocr_unmatched and not self.tab4_stop_flag:
                self.log(f"有 {len(llm_ocr_unmatched)} 个文件OCR未匹配，移动到'0未处理文件夹'目录")
                
                unprocessed_dir = os.path.join(dest_dir, "0未处理文件夹")
                os.makedirs(unprocessed_dir, exist_ok=True)
                
                unprocessed_records = []  # 用于生成Excel
                
                for file_path, ocr_text in llm_ocr_unmatched:
                    file_name = os.path.basename(file_path)
                    dest_file = os.path.join(unprocessed_dir, file_name)
                    
                    # 处理文件名冲突
                    counter = 1
                    original_dest = dest_file
                    while os.path.exists(dest_file):
                        name, ext = os.path.splitext(original_dest)
                        dest_file = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    shutil.move(file_path, dest_file)
                    
                    # 记录文件名和OCR内容
                    unprocessed_records.append({
                        "文件名": file_name,
                        "OCR识别内容": ocr_text if ocr_text else "(无识别结果)"
                    })
                
                # 生成Excel记录文件（放到0-OCR识别结果文件夹）
                if unprocessed_records:
                    try:
                        ocr_result_dir = os.path.join(dest_dir, "0-OCR识别结果")
                        os.makedirs(ocr_result_dir, exist_ok=True)
                        excel_path = os.path.join(ocr_result_dir, "LLMOCR识别记录.xlsx")
                        df_unprocessed = pd.DataFrame(unprocessed_records)
                        df_unprocessed.to_excel(excel_path, index=False)
                        self.log(f"✓ 已生成LLMOCR识别记录Excel: {excel_path}")
                    except Exception as e:
                        self.log(f"⚠ 生成Excel记录失败: {e}")
                
                # 清理空的OCR待处理目录（使用shutil.rmtree确保完全删除）
                if os.path.exists(ocr_pending_dir):
                    try:
                        # 检查目录是否为空或只包含空子目录
                        has_content = False
                        for root, dirs, files in os.walk(ocr_pending_dir):
                            if files:
                                has_content = True
                                break
                        
                        if not has_content:
                            shutil.rmtree(ocr_pending_dir)
                            self.log("OCR待处理文件夹已空，删除该目录")
                        else:
                            self.log("⚠ OCR待处理文件夹中仍有文件，保留该目录")
                    except Exception as e:
                        self.log(f"⚠ 清理OCR待处理文件夹时出错: {e}")
                
                # 如果有可选关键列，对未处理文件进行再次匹配
                if key_col_opt:
                    self.log(f"使用可选关键列'{key_col_opt}'对未处理文件进行二次匹配...")
                    
                    # 构建可选关键列的映射
                    opt_data_map = {}
                    opt_keywords = []
                    for _, row in self.tab4_df.iterrows():
                        opt_val = str(row.get(key_col_opt, "")).strip()
                        if opt_val and opt_val != 'nan':
                            opt_keywords.append(opt_val)
                            opt_data_map[opt_val] = row
                    
                    # 遍历未处理目录中的文件
                    unprocessed_files = [f for f in os.listdir(unprocessed_dir) 
                                         if os.path.splitext(f)[1].lower() in valid_exts]
                    
                    for file_name in unprocessed_files:
                        if self.tab4_stop_flag:
                            break
                        
                        file_path = os.path.join(unprocessed_dir, file_name)
                        matched = False
                        
                        # 文件名匹配
                        for keyword in opt_keywords:
                            patterns = [f"*{keyword}*", f"{keyword}_*", f"*_{keyword}", f"{keyword}*", f"*{keyword}"]
                            for pattern in patterns:
                                if fnmatch.fnmatch(file_name, pattern):
                                    row = opt_data_map[keyword]
                                    file_title = self._extract_file_title(file_path, title_page) if "[文件标题]" in folder_cols or "[文件标题]" in file_cols else ""
                                    folder_name = self._build_folder_name(row, folder_cols, file_title)
                                    new_file_name = build_file_name(row, file_cols, file_name, file_title)
                                    
                                    target_folder = os.path.join(dest_dir, folder_name)
                                    os.makedirs(target_folder, exist_ok=True)
                                    
                                    dest_file = os.path.join(target_folder, new_file_name)
                                    shutil.move(file_path, dest_file)
                                    self.log(f"✓ 可选列匹配归档: {file_name} -> {folder_name}/{new_file_name}")
                                    matched = True
                                    break
                            if matched:
                                break
                    
                    # 清理空的未处理目录
                    remaining_files = os.listdir(unprocessed_dir)
                    # 排除Excel文件检查
                    non_excel_files = [f for f in remaining_files if not f.endswith('.xlsx')]
                    if not non_excel_files:
                        # 删除Excel文件后再删除目录
                        for f in remaining_files:
                            try:
                                os.remove(os.path.join(unprocessed_dir, f))
                            except:
                                pass
                        try:
                            os.rmdir(unprocessed_dir)
                            self.log("所有文件已处理，删除'0未处理文件夹'目录")
                        except:
                            pass
                
            if not self.tab4_stop_flag:
                # 打印Token消耗汇总
                if token_usage_total['total_tokens'] > 0:
                    self.log(f"📊📊📊 本次任务Token消耗汇总 📊📊📊")
                    self.log(f"📊 输入Token: {token_usage_total['prompt_tokens']}")
                    self.log(f"📊 输出Token: {token_usage_total['completion_tokens']}")
                    self.log(f"📊 总计Token: {token_usage_total['total_tokens']}")
                self.log(f"归档任务完成！共处理 {len(processed_files)} 个文件")
                self.root.after(0, lambda: messagebox.showinfo("完成", f"归档任务完成！\n共处理 {len(processed_files)} 个文件"))
            else:
                # 打印Token消耗汇总（即使任务被停止）
                if token_usage_total['total_tokens'] > 0:
                    self.log(f"📊📊📊 本次任务Token消耗汇总（任务已停止） 📊📊📊")
                    self.log(f"📊 输入Token: {token_usage_total['prompt_tokens']}")
                    self.log(f"📊 输出Token: {token_usage_total['completion_tokens']}")
                    self.log(f"📊 总计Token: {token_usage_total['total_tokens']}")
                self.log(f"任务已停止，已处理 {len(processed_files)} 个文件")
                
        except Exception as e:
            self.log(f"归档任务错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"归档任务失败:\n{e}"))
        finally:
            self.update_tab4_ui_state(False)

    def setup_tab5(self):
        """Tab5: 向量智能归档 - 基于向量模型的文件自动归档"""
        # Tab5变量初始化
        self.tab5_excel_path_var = tk.StringVar()
        # 2个绝对关键列 + 3个辅助关键列
        self.tab5_abs_key1_var = tk.StringVar()
        self.tab5_abs_key2_var = tk.StringVar()
        self.tab5_aux_key1_var = tk.StringVar()
        self.tab5_aux_key2_var = tk.StringVar()
        self.tab5_aux_key3_var = tk.StringVar()
        # OCR配置（复用Tab4模式）
        self.tab5_ocr_pages_var = tk.StringVar(value="1")
        self.tab5_ocr_max_chars_var = tk.StringVar(value="")
        self.tab5_ocr_prompt_var = tk.StringVar(value="")
        self.tab5_skip_hash_check_var = tk.BooleanVar(value=False)
        # 向量模型配置（Qwen）
        self.tab5_vector_api_key_var = tk.StringVar(value=self.config.get("tab5_vector_api_key", ""))
        self.tab5_vector_model_var = tk.StringVar(value=self.config.get("tab5_vector_model", "text-embedding-v3"))
        # 火山引擎OCR配置（可复用Tab4或独立配置）
        self.tab5_use_tab4_ocr_var = tk.BooleanVar(value=self.config.get("tab5_use_tab4_ocr", True))
        self.tab5_ocr_api_key_var = tk.StringVar(value=self.config.get("tab5_ocr_api_key", ""))
        self.tab5_ocr_model_var = tk.StringVar(value=self.config.get("tab5_ocr_model", "doubao-seed-1-6-vision-250815"))
        # 目录配置
        self.tab5_source_dir_var = tk.StringVar()
        self.tab5_dest_dir_var = tk.StringVar()
        # 冲突处理
        self.tab5_conflict_auto_resolve_var = tk.BooleanVar(value=True)
        self.tab5_conflict_dir_var = tk.StringVar(value="冲突待处理")
        
        # Tab5 文件夹命名配置（独立于Tab4）
        self.tab5_folder_name_col1_var = tk.StringVar(value=self.config.get("tab5_folder_name_col1", ""))
        self.tab5_folder_name_col2_var = tk.StringVar(value=self.config.get("tab5_folder_name_col2", ""))
        self.tab5_folder_name_col3_var = tk.StringVar(value=self.config.get("tab5_folder_name_col3", ""))
        
        # Tab5 文件名命名配置（独立于Tab4）
        self.tab5_file_name_col1_var = tk.StringVar(value=self.config.get("tab5_file_name_col1", ""))
        self.tab5_file_name_col2_var = tk.StringVar(value=self.config.get("tab5_file_name_col2", ""))
        self.tab5_file_name_col3_var = tk.StringVar(value=self.config.get("tab5_file_name_col3", ""))
        # 状态变量
        self.tab5_df = None
        self.tab5_is_running = False
        self.tab5_stop_flag = False
        
        # Tab6 变量初始化
        self.tab6_template_file_var = tk.StringVar()
        self.tab6_data_file_var = tk.StringVar()
        self.tab6_output_dir_var = tk.StringVar()
        self.tab6_convert_to_pdf_var = tk.BooleanVar(value=False)
        self.tab6_filename_template_var = tk.StringVar(value="{第一列}")
        self.tab6_is_running = False
        self.tab6_stop_flag = False
        self.tab6_progress_var = tk.DoubleVar()
        self.tab6_status_var = tk.StringVar(value="待开始")
        
        frame = self.tab5
        frame.configure(bg=COLORS['bg_primary'])
        
        # 创建滚动容器
        main_canvas = tk.Canvas(frame, bg=COLORS['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=COLORS['bg_primary'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        self.scrollable_window_tab5 = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        def on_canvas_configure(event):
            canvas_width = event.width - 4
            main_canvas.itemconfig(self.scrollable_window_tab5, width=canvas_width)
        main_canvas.bind('<Configure>', on_canvas_configure)
        
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            main_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            main_canvas.unbind_all("<MouseWheel>")
        
        main_canvas.bind('<Enter>', bind_mousewheel)
        main_canvas.bind('<Leave>', unbind_mousewheel)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 标题
        title_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="向量智能归档",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="基于向量模型语义匹配，智能识别并归档文件",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))
        
        # 1. Excel数据源卡片
        excel_card = self.create_card(scrollable_frame, "Excel数据源")
        excel_card.pack(fill="x", pady=8)
        
        # Excel文件选择
        excel_row = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        excel_row.pack(fill="x", pady=(0, 12))
        tk.Entry(excel_row, 
                textvariable=self.tab5_excel_path_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.create_secondary_button(excel_row, "加载Excel", self.load_excel_tab5).pack(side="right")
        
        # 列选择区域
        col_grid = tk.Frame(excel_card, bg=COLORS['bg_primary'])
        col_grid.pack(fill="x")
        
        # 绝对关键列1（必填，优先级最高）
        tk.Label(col_grid, 
                text="绝对关键列1*:",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold')).grid(row=0, column=0, sticky="e", pady=6)
        self.tab5_abs_key1_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab5_abs_key1_cb.grid(row=0, column=1, padx=(8, 16), sticky="w")
        
        # 绝对关键列2（必填，OR逻辑）
        tk.Label(col_grid, 
                text="绝对关键列2*:",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold')).grid(row=0, column=2, sticky="e", pady=6)
        self.tab5_abs_key2_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab5_abs_key2_cb.grid(row=0, column=3, padx=(8, 0), sticky="w")
        
        # 辅助关键列
        aux_label = tk.Label(col_grid, 
                text="【辅助关键列】(可选，用于提高匹配精度)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9, 'bold'))
        aux_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(12, 4))
        
        # 辅助列1
        tk.Label(col_grid, 
                text="辅助列1:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=0, sticky="e", pady=6)
        self.tab5_aux_key1_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab5_aux_key1_cb.grid(row=2, column=1, padx=(8, 16), sticky="w")
        
        # 辅助列2
        tk.Label(col_grid, 
                text="辅助列2:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=2, column=2, sticky="e", pady=6)
        self.tab5_aux_key2_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab5_aux_key2_cb.grid(row=2, column=3, padx=(8, 0), sticky="w")
        
        # 辅助列3
        tk.Label(col_grid, 
                text="辅助列3:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=3, column=0, sticky="e", pady=6)
        self.tab5_aux_key3_cb = ttk.Combobox(col_grid, state="readonly", width=15)
        self.tab5_aux_key3_cb.grid(row=3, column=1, padx=(8, 16), sticky="w")
        
        # 关键列说明
        hint_label = tk.Label(col_grid, 
                text="说明: 只需命中任一绝对关键列即可归档，列1优先级高于列2",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 8))
        hint_label.grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))
        
        # 2. 向量模型配置卡片
        vector_card = self.create_card(scrollable_frame, "向量模型配置 (Qwen)")
        vector_card.pack(fill="x", pady=8)
        
        vector_grid = tk.Frame(vector_card, bg=COLORS['bg_primary'])
        vector_grid.pack(fill="x")
        
        # API Key
        tk.Label(vector_grid, 
                text="API Key:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(vector_grid, 
                textvariable=self.tab5_vector_api_key_var, 
                show="●",
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        self.create_secondary_button(vector_grid, "测试连接", self.test_vector_connection_tab5).grid(row=0, column=2)
        
        # Model ID（固定）
        tk.Label(vector_grid, 
                text="模型ID:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Label(vector_grid, 
                text="text-embedding-v3 (1024维)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 9)).grid(row=1, column=1, padx=(12, 8), sticky="w")
        
        # 3. OCR配置卡片
        ocr_card = self.create_card(scrollable_frame, "OCR配置")
        ocr_card.pack(fill="x", pady=8)
        
        ocr_grid = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        ocr_grid.pack(fill="x")
        
        # OCR页码配置
        tk.Label(ocr_grid, 
                text="OCR识别页码:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab5_ocr_pages_var,
                width=15,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(支持多页码，如: 1,2,3 或 1-3,5)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="w")
        
        # 识别文字量配置
        tk.Label(ocr_grid, 
                text="识别文字量(可选):",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Entry(ocr_grid, 
                textvariable=self.tab5_ocr_max_chars_var,
                width=10,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=1, column=1, padx=(12, 8), sticky="w")
        tk.Label(ocr_grid, 
                text="(LLM OCR时最大输出字数)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=2, sticky="w")
        
                
        # 独立OCR配置（当不复用时显示）
        ocr_config_frame = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        ocr_config_frame.pack(fill="x", pady=(8, 0))
        
        tk.Label(ocr_config_frame, 
                text="独立OCR配置:",
                bg=COLORS['bg_primary'],
                fg=COLORS['secondary'],
                font=('微软雅黑', 9, 'bold')).pack(anchor="w", pady=(0, 8))
        
        ocr_ai_grid = tk.Frame(ocr_config_frame, bg=COLORS['bg_primary'])
        ocr_ai_grid.pack(fill="x")
        
        tk.Label(ocr_ai_grid, 
                text="API Key:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", pady=6)
        tk.Entry(ocr_ai_grid, 
                textvariable=self.tab5_ocr_api_key_var, 
                show="●",
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=0, column=1, padx=(12, 8), sticky="w")
        
        tk.Label(ocr_ai_grid, 
                text="模型ID:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=1, column=0, sticky="e", pady=6)
        tk.Entry(ocr_ai_grid, 
                textvariable=self.tab5_ocr_model_var, 
                width=35,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).grid(row=1, column=1, padx=(12, 8), sticky="w")
        tk.Button(ocr_ai_grid, text="测试连接", 
                  command=self.test_ai_connection_tab5,
                  bg=COLORS['input_bg'], fg=COLORS['text_primary'],
                  relief="solid", bd=1, padx=10).grid(row=0, column=2, rowspan=2, padx=10)
        
        # Hash缓存选项
        hash_frame = tk.Frame(ocr_card, bg=COLORS['bg_primary'])
        hash_frame.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(hash_frame, 
                      text="跳过Hash核对（强制重新OCR）",
                      variable=self.tab5_skip_hash_check_var,
                      bg=COLORS['bg_primary'],
                      fg=COLORS['text_primary'],
                      selectcolor=COLORS['bg_primary'],
                      font=('微软雅黑', 9)).pack(anchor="w")
        
        # 4. 文件夹配置卡片
        folder_card = self.create_card(scrollable_frame, "文件夹配置")
        folder_card.pack(fill="x", pady=8)
        
        # 文件源
        src_row = tk.Frame(folder_card, bg=COLORS['bg_primary'])
        src_row.pack(fill="x", pady=(0, 8))
        tk.Label(src_row, 
                text="文件源:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(src_row, 
                textvariable=self.tab5_source_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8), fill="x", expand=True)
        self.create_secondary_button(src_row, "浏览...", lambda: self.select_dir(self.tab5_source_dir_var)).pack(side="right")
        
        # 归档目的地
        dest_row = tk.Frame(folder_card, bg=COLORS['bg_primary'])
        dest_row.pack(fill="x")
        tk.Label(dest_row, 
                text="归档目的地:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(dest_row, 
                textvariable=self.tab5_dest_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8), fill="x", expand=True)
        self.create_secondary_button(dest_row, "浏览...", lambda: self.select_dir(self.tab5_dest_dir_var)).pack(side="right")
        
        # 5. 命名规则卡片
        naming_card = self.create_card(scrollable_frame, "命名规则")
        naming_card.pack(fill="x", pady=8)
        
        # 文件夹命名
        folder_naming_frame = tk.Frame(naming_card, bg=COLORS['bg_primary'])
        folder_naming_frame.pack(fill="x", pady=(0, 12))
        
        folder_label = tk.Label(folder_naming_frame, 
                text="【文件夹命名】(按顺序组合)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 9, 'bold'))
        folder_label.pack(anchor="w", pady=(0, 6))
        
        folder_cols_row = tk.Frame(folder_naming_frame, bg=COLORS['bg_primary'])
        folder_cols_row.pack(fill="x")
        
        # 文件夹列1
        tk.Label(folder_cols_row, 
                text="列1:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", padx=(0, 4), pady=4)
        self.tab5_folder_name_col1_cb = ttk.Combobox(folder_cols_row, state="readonly", width=20)
        self.tab5_folder_name_col1_cb.grid(row=0, column=1, padx=(4, 16), sticky="w")
        
        # 文件夹列2
        tk.Label(folder_cols_row, 
                text="列2:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="e", padx=(0, 4), pady=4)
        self.tab5_folder_name_col2_cb = ttk.Combobox(folder_cols_row, state="readonly", width=20)
        self.tab5_folder_name_col2_cb.grid(row=0, column=3, padx=(4, 16), sticky="w")
        
        # 文件夹列3
        tk.Label(folder_cols_row, 
                text="列3:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=4, sticky="e", padx=(0, 4), pady=4)
        self.tab5_folder_name_col3_cb = ttk.Combobox(folder_cols_row, state="readonly", width=20)
        self.tab5_folder_name_col3_cb.grid(row=0, column=5, padx=(4, 0), sticky="w")
        
        # 文件命名
        file_naming_frame = tk.Frame(naming_card, bg=COLORS['bg_primary'])
        file_naming_frame.pack(fill="x")
        
        file_label = tk.Label(file_naming_frame, 
                text="【文件命名】(按顺序组合)",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 9, 'bold'))
        file_label.pack(anchor="w", pady=(0, 6))
        
        file_cols_row = tk.Frame(file_naming_frame, bg=COLORS['bg_primary'])
        file_cols_row.pack(fill="x")
        
        # 文件列1
        tk.Label(file_cols_row, 
                text="列1:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=0, sticky="e", padx=(0, 4), pady=4)
        self.tab5_file_name_col1_cb = ttk.Combobox(file_cols_row, state="readonly", width=20)
        self.tab5_file_name_col1_cb.grid(row=0, column=1, padx=(4, 16), sticky="w")
        
        # 文件列2
        tk.Label(file_cols_row, 
                text="列2:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=2, sticky="e", padx=(0, 4), pady=4)
        self.tab5_file_name_col2_cb = ttk.Combobox(file_cols_row, state="readonly", width=20)
        self.tab5_file_name_col2_cb.grid(row=0, column=3, padx=(4, 16), sticky="w")
        
        # 文件列3
        tk.Label(file_cols_row, 
                text="列3:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).grid(row=0, column=4, sticky="e", padx=(0, 4), pady=4)
        self.tab5_file_name_col3_cb = ttk.Combobox(file_cols_row, state="readonly", width=20)
        self.tab5_file_name_col3_cb.grid(row=0, column=5, padx=(4, 0), sticky="w")
        
        # 6. 冲突处理选项卡片
        conflict_card = self.create_card(scrollable_frame, "冲突处理")
        conflict_card.pack(fill="x", pady=8)
        
        conflict_frame = tk.Frame(conflict_card, bg=COLORS['bg_primary'])
        conflict_frame.pack(fill="x")
        
        tk.Checkbutton(conflict_frame, 
                      text="自动选择权重最高的匹配（否则暂停等待确认）",
                      variable=self.tab5_conflict_auto_resolve_var,
                      bg=COLORS['bg_primary'],
                      fg=COLORS['text_primary'],
                      selectcolor=COLORS['bg_primary'],
                      font=('微软雅黑', 9)).pack(anchor="w")
        
        conflict_dir_row = tk.Frame(conflict_card, bg=COLORS['bg_primary'])
        conflict_dir_row.pack(fill="x", pady=(8, 0))
        tk.Label(conflict_dir_row, 
                text="冲突文件暂存目录:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(side="left")
        tk.Entry(conflict_dir_row, 
                textvariable=self.tab5_conflict_dir_var,
                width=20,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(12, 8))
        
        # 6. 执行按钮区域
        action_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'], pady=20)
        action_frame.pack(fill="x", side="bottom")
        
        button_container = tk.Frame(action_frame, bg=COLORS['bg_primary'])
        button_container.pack(fill="x")
        
        self.tab5_start_btn = self.create_primary_button(button_container, "▶ 开始向量归档", self.start_vector_archive_task)
        self.tab5_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.tab5_stop_btn = tk.Button(button_container, 
                                     text="■ 停止任务", 
                                     command=self.stop_vector_archive_task,
                                     bg=COLORS['danger'],
                                     fg='white',
                                     font=('微软雅黑', 10, 'bold'),
                                     relief='flat',
                                     padx=20,
                                     pady=8,
                                     cursor='hand2',
                                     state="disabled")
        self.tab5_stop_btn.pack(side="right", fill="x", expand=True)

    def load_excel_tab5(self):
        """加载Excel文件并更新Tab5列选择器"""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            try:
                self.tab5_excel_path_var.set(path)
                self.tab5_df = pd.read_excel(path)
                cols = self.tab5_df.columns.tolist()
                cols_with_empty = [""] + cols
                
                # 更新所有关键列选择器
                self.tab5_abs_key1_cb['values'] = cols
                self.tab5_abs_key2_cb['values'] = cols
                self.tab5_aux_key1_cb['values'] = cols_with_empty
                self.tab5_aux_key2_cb['values'] = cols_with_empty
                self.tab5_aux_key3_cb['values'] = cols_with_empty
                
                # 更新文件夹和文件命名选择器
                self.tab5_folder_name_col1_cb['values'] = cols_with_empty
                self.tab5_folder_name_col2_cb['values'] = cols_with_empty
                self.tab5_folder_name_col3_cb['values'] = cols_with_empty
                self.tab5_file_name_col1_cb['values'] = cols_with_empty
                self.tab5_file_name_col2_cb['values'] = cols_with_empty
                self.tab5_file_name_col3_cb['values'] = cols_with_empty
                
                self.log(f"Tab5 Excel加载成功: {len(self.tab5_df)} 行, {len(cols)} 列")
            except Exception as e:
                messagebox.showerror("错误", f"Excel读取失败: {e}")
                self.tab5_df = None

    def setup_tab6(self):
        """设置Tab6: 文件批量填充"""
        frame = self.tab6
        frame.configure(bg=COLORS['bg_primary'])
        
        # 主容器 - 使用grid布局实现更好的响应式设计
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        # 创建滚动容器 - 支持鼠标滚轮
        main_canvas = tk.Canvas(frame, bg=COLORS['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg=COLORS['bg_primary'])
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            main_canvas.unbind_all("<MouseWheel>")
        
        main_canvas.bind('<Enter>', _bind_to_mousewheel)
        main_canvas.bind('<Leave>', _unbind_from_mousewheel)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        self.scrollable_window_tab6 = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # 标题
        title_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 16))
        tk.Label(title_frame, 
                text="文件批量填充",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 16, 'bold')).pack(anchor="w")
        tk.Label(title_frame,
                text="基于Word模板和Excel数据，批量生成文档",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9)).pack(anchor="w", pady=(4, 0))
        
        # 1. 文件选择卡片
        file_card = self.create_card(scrollable_frame, "文件选择")
        file_card.pack(fill="x", pady=8)
        
        # 模板文件选择
        template_row = tk.Frame(file_card, bg=COLORS['bg_primary'])
        template_row.pack(fill="x", pady=(0, 8))
        tk.Label(template_row, 
                text="Word模板:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9), width=12, anchor="e").pack(side="left")
        tk.Entry(template_row, 
                textvariable=self.tab6_template_file_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(8, 8), fill="x", expand=True)
        self.create_secondary_button(template_row, "浏览...", self.select_template_file).pack(side="right")
        
        # 数据文件选择
        data_row = tk.Frame(file_card, bg=COLORS['bg_primary'])
        data_row.pack(fill="x", pady=(0, 8))
        tk.Label(data_row, 
                text="Excel数据:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9), width=12, anchor="e").pack(side="left")
        tk.Entry(data_row, 
                textvariable=self.tab6_data_file_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(8, 8), fill="x", expand=True)
        self.create_secondary_button(data_row, "浏览...", self.select_data_file).pack(side="right")
        
        # 输出目录选择
        output_row = tk.Frame(file_card, bg=COLORS['bg_primary'])
        output_row.pack(fill="x")
        tk.Label(output_row, 
                text="输出目录:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9), width=12, anchor="e").pack(side="left")
        tk.Entry(output_row, 
                textvariable=self.tab6_output_dir_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(8, 8), fill="x", expand=True)
        self.create_secondary_button(output_row, "浏览...", self.select_output_dir).pack(side="right")
        
        # 2. 命名和打印选项卡片
        option_card = self.create_card(scrollable_frame, "命名和打印选项")
        option_card.pack(fill="x", pady=8)
        
        # 文件名模板
        filename_row = tk.Frame(option_card, bg=COLORS['bg_primary'])
        filename_row.pack(fill="x", pady=(0, 8))
        tk.Label(filename_row, 
                text="文件名模板:",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_secondary'],
                font=('微软雅黑', 9), width=12, anchor="e").pack(side="left")
        tk.Entry(filename_row, 
                textvariable=self.tab6_filename_template_var,
                bg=COLORS['input_bg'],
                fg=COLORS['text_primary'],
                insertbackground=COLORS['text_primary'],
                relief='solid',
                highlightthickness=1,
                highlightcolor=COLORS['border']).pack(side="left", padx=(8, 8), fill="x", expand=True)
        tk.Label(filename_row, 
                text="例: {姓名}_{部门}_合同",
                bg=COLORS['bg_primary'],
                fg=COLORS['text_tertiary'],
                font=('微软雅黑', 8)).pack(side="right")
        
        # PDF转换选项
        pdf_frame = tk.Frame(option_card, bg=COLORS['bg_primary'])
        pdf_frame.pack(fill="x")
        
        tk.Checkbutton(pdf_frame, 
                      text="转换为PDF（使用Word原生导出功能，100%静默）",
                      variable=self.tab6_convert_to_pdf_var,
                      bg=COLORS['bg_primary'],
                      fg=COLORS['text_primary'],
                      selectcolor=COLORS['bg_primary'],
                      font=('微软雅黑', 9)).pack(anchor="w")
        
        # 3. 执行按钮区域
        action_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_primary'], pady=20)
        action_frame.pack(fill="x")
        
        button_container = tk.Frame(action_frame, bg=COLORS['bg_primary'])
        button_container.pack(fill="x")
        
        self.tab6_start_btn = self.create_primary_button(button_container, "▶ 开始批量填充", self.start_batch_fill)
        self.tab6_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.tab6_stop_btn = tk.Button(button_container, 
                                     text="■ 停止任务", 
                                     command=self.stop_batch_fill,
                                     bg=COLORS['danger'],
                                     fg='white',
                                     font=('微软雅黑', 10, 'bold'),
                                     relief='flat',
                                     padx=20,
                                     pady=8,
                                     cursor='hand2',
                                     state="disabled")
        self.tab6_stop_btn.pack(side="right", fill="x", expand=True)
        
        # 4. 固定进度条在底部（不跟随滚动）
        progress_container = tk.Frame(frame, bg=COLORS['bg_secondary'], relief='ridge', bd=1)
        progress_container.grid(row=1, column=0, columnspan=2, sticky="ew")
        
        progress_inner = tk.Frame(progress_container, bg=COLORS['bg_secondary'], padx=16, pady=8)
        progress_inner.pack(fill="x")
        
        tk.Label(progress_inner, 
                textvariable=self.tab6_status_var,
                bg=COLORS['bg_secondary'],
                fg=COLORS['text_primary'],
                font=('微软雅黑', 9)).pack(anchor="w")
        
        self.tab6_progress_bar = ttk.Progressbar(progress_inner, 
                                               variable=self.tab6_progress_var,
                                               maximum=100,
                                               mode='determinate')
        self.tab6_progress_bar.pack(fill="x", pady=(4, 0))

    def select_template_file(self):
        """选择Word模板文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.tab6_template_file_var.set(file_path)
    
    def select_data_file(self):
        """选择Excel数据文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.tab6_data_file_var.set(file_path)
    
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.tab6_output_dir_var.set(dir_path)
    

    def start_batch_fill(self):
        """启动批量填充任务"""
        # 参数校验
        template_file = self.tab6_template_file_var.get().strip()
        data_file = self.tab6_data_file_var.get().strip()
        output_dir = self.tab6_output_dir_var.get().strip()
        
        if not template_file or not data_file or not output_dir:
            messagebox.showwarning("配置不全", "请选择模板文件、数据文件和输出目录")
            return
        
        if not os.path.exists(template_file):
            messagebox.showerror("错误", "模板文件不存在")
            return
        
        if not os.path.exists(data_file):
            messagebox.showerror("错误", "数据文件不存在")
            return
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        if self.tab6_convert_to_pdf_var.get():
            pdf_dir = os.path.join(output_dir, "PDF文件")
            os.makedirs(pdf_dir, exist_ok=True)
        
        # 重置状态
        self.tab6_stop_flag = False
        self.tab6_progress_var.set(0)
        self.tab6_status_var.set("准备中...")
        
        # 启动异步任务
        threading.Thread(target=self._run_batch_fill_task, args=(template_file, data_file, output_dir), daemon=True).start()
        
        # 更新UI状态
        self.tab6_is_running = True
        self.tab6_start_btn.configure(state="disabled")
        self.tab6_stop_btn.configure(state="normal")
    
    def stop_batch_fill(self):
        """停止批量填充任务"""
        if self.tab6_is_running:
            self.tab6_stop_flag = True
            self.log("🛑 正在停止批量填充任务...")
            self.tab6_stop_btn.configure(state="disabled", text="停止中...")
    
    def _run_batch_fill_task(self, template_file, data_file, output_dir):
        """执行批量填充任务（异步）"""
        try:
            import gc
            import io
            import pandas as pd
            from docxtpl import DocxTemplate
            import jinja2
            from tqdm import tqdm
            import win32com.client
            import pythoncom
            import time
        except ImportError as e:
            self.root.after(0, lambda: messagebox.showerror("缺少依赖", f"请安装必要的库: {e}"))
            return
        
        # 保护机制：如果 Word 模板中存在其他大括号 {}，但不属于变量，则保留原样
        class KeepUndefined(jinja2.Undefined):
            def __str__(self):
                return f"{{{self._undefined_name}}}"
        
        try:
            self.root.after(0, lambda: self.tab6_status_var.set("读取数据文件..."))
            
            # 读取Excel数据
            df = pd.read_excel(data_file, dtype=str, keep_default_na=False)
            total_rows = len(df)
            
            if total_rows == 0:
                self.root.after(0, lambda: messagebox.showwarning("警告", "数据文件中没有数据"))
                return
            
            # 读取模板文件
            with open(template_file, "rb") as f:
                template_bytes = f.read()
            
            # 获取列信息
            all_columns = df.columns.tolist()
            placeholder_cols = [col for col in all_columns if '{' in col and '}' in col]
            
            if not placeholder_cols:
                self.root.after(0, lambda: messagebox.showwarning("警告", "在Excel的表头中没有找到被{}囊括的列名"))
                return
            
            # 获取文件名模板
            filename_template = self.tab6_filename_template_var.get().strip()
            if not filename_template:
                filename_template = "{第一列}"
            
            # 准备进度计算
            convert_to_pdf = self.tab6_convert_to_pdf_var.get()
            if convert_to_pdf:
                # 如果需要转换PDF，总任务量包括生成Word和转换PDF
                total_tasks = total_rows * 2  # 每行两个任务：生成Word + 转换PDF
            else:
                total_tasks = total_rows  # 只需要生成Word
            
            completed_tasks = 0
            
            # 设置Jinja2环境
            myenv = jinja2.Environment(
                variable_start_string='{', 
                variable_end_string='}',
                undefined=KeepUndefined
            )
            
            self.root.after(0, lambda: self.log(f"开始批量生成 {total_rows} 个文件..."))
            
            # 批量生成Word文件
            processed_count = 0
            error_files = []
            
            for index, row in df.iterrows():
                if self.tab6_stop_flag:
                    break
                
                try:
                    # 生成文件名 - 使用模板
                    filename_dict = {}
                    
                    # 首先处理模板中的变量（带{}的格式）
                    import re
                    template_vars = re.findall(r'\{([^}]+)\}', filename_template)
                    
                    # 为每个变量准备值
                    for var in template_vars:
                        # 查找Excel中是否有对应的列
                        matching_col = None
                        for col in all_columns:
                            # 去掉列名中的{}进行匹配
                            clean_col = col.replace('{', '').replace('}', '').strip()
                            if clean_col == var.strip():
                                matching_col = col
                                break
                        
                        if matching_col:
                            filename_dict[var] = str(row[matching_col]).strip()
                        else:
                            # 如果没找到匹配的列，使用变量名本身
                            filename_dict[var] = var
                            warning_msg = f"  警告: 未找到列 '{var}'，将使用变量名"
                            self.root.after(0, lambda msg=warning_msg: self.log(msg))
                    
                    # 替换模板中的变量
                    try:
                        raw_file_name = filename_template.format(**filename_dict)
                    except KeyError as e:
                        # 如果还有未处理的变量，使用第一列作为后备
                        warning_msg = "  警告: 文件名模板格式错误，使用第一列作为文件名"
                        self.root.after(0, lambda msg=warning_msg: self.log(msg))
                        raw_file_name = str(row[all_columns[0]]).strip()
                    
                    # 清理文件名中的非法字符
                    for char in '<>:"/\\|?*':
                        raw_file_name = raw_file_name.replace(char, "_")
                    
                    new_file_name = f"{raw_file_name}.docx"
                    new_file_path = os.path.join(output_dir, new_file_name)
                    
                    # 准备上下文数据
                    context = {}
                    for col in placeholder_cols:
                        var_name = col.replace('{', '').replace('}', '').strip()
                        context[var_name] = str(row[col])
                    
                    # 生成Word文档
                    with io.BytesIO(template_bytes) as stream:
                        doc = DocxTemplate(stream)
                        doc.render(context, jinja_env=myenv)
                        doc.save(new_file_path)
                        del doc
                    
                    processed_count += 1
                    completed_tasks += 1
                    
                    # 更新进度
                    progress = completed_tasks / total_tasks * 100
                    self.root.after(0, lambda p=progress, s=f"生成Word... {processed_count}/{total_rows}": (
                        self.tab6_progress_var.set(p),
                        self.tab6_status_var.set(s)
                    ))
                    
                    # 内存回收
                    if (index + 1) % 50 == 0:
                        gc.collect()
                        time.sleep(0.1)  # 防止CPU占用过高
                    
                except Exception as e:
                    error_files.append(f"第{index+1}行: {str(e)}")
                    error_msg = f"  × 处理第{index+1}行失败: {e}"
                    self.root.after(0, lambda msg=error_msg: self.log(msg))
            
            # 如果需要转换为PDF
            if convert_to_pdf and not self.tab6_stop_flag:
                self.root.after(0, lambda: self.tab6_status_var.set("转换为PDF中..."))
                self.root.after(0, lambda: self.log("开始转换为PDF..."))
                
                # 使用 aspose-words 进行 PDF 转换（避免 Word 进程）
                try:
                    import aspose.words as aw
                    
                    pdf_dir = os.path.join(output_dir, "PDF文件")
                    os.makedirs(pdf_dir, exist_ok=True)
                    pdf_errors = []
                    
                    self.root.after(0, lambda: self.log("使用 Aspose.Words 进行 PDF 转换（无 Word 进程）..."))
                    
                    # 获取所有 Word 文件
                    word_files = [f for f in os.listdir(output_dir) if f.endswith('.docx')]
                    
                    # 批量转换
                    for i, docx_file in enumerate(word_files):
                        if self.tab6_stop_flag:
                            break
                        
                        try:
                            docx_path = os.path.join(output_dir, docx_file)
                            pdf_path = os.path.join(pdf_dir, docx_file.replace('.docx', '.pdf'))
                            
                            start_time = time.time()
                            
                            # 使用 Aspose.Words 转换（纯内存操作，无外部进程）
                            doc = aw.Document(docx_path)
                            doc.save(pdf_path)
                            
                            file_size = os.path.getsize(pdf_path)
                            conversion_time = time.time() - start_time
                            
                            success_msg = f"PDF转换成功: {docx_file} (耗时: {conversion_time:.2f}s, 大小: {file_size//1024}KB)"
                            self.root.after(0, lambda msg=success_msg: self.log(msg))
                            
                            completed_tasks += 1
                            
                            # 更新进度
                            progress = completed_tasks / total_tasks * 100
                            self.root.after(0, lambda p=progress, s=f"PDF转换中... {i + 1}/{len(word_files)}": (
                                self.tab6_progress_var.set(p),
                                self.tab6_status_var.set(s)
                            ))
                            
                            # 内存回收
                            del doc
                            if (i + 1) % 10 == 0:
                                gc.collect()
                            
                        except Exception as e:
                            pdf_errors.append(f"{docx_file}: {str(e)}")
                            pdf_error_msg = f"  × PDF转换失败 {docx_file}: {e}"
                            self.root.after(0, lambda msg=pdf_error_msg: self.log(msg))
                            
                except ImportError as e:
                    error_msg = f"缺少 aspose-words 库。请运行: pip install aspose-words"
                    self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                    self.root.after(0, lambda msg=error_msg: self.log(msg))
                    return
                except Exception as e:
                    error_msg = f"PDF转换初始化失败: {e}"
                    self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                    self.root.after(0, lambda msg=error_msg: self.log(msg))
                    return
            
            # 任务完成
            if not self.tab6_stop_flag:
                self.root.after(0, lambda: self.tab6_status_var.set("完成"))
                self.root.after(0, lambda: self.tab6_progress_var.set(100))
                
                # 显示结果
                result_msg = f"批量填充完成！\n\n成功处理: {processed_count} 个文件"
                if error_files:
                    result_msg += f"\n错误文件: {len(error_files)} 个"
                
                if self.tab6_convert_to_pdf_var.get():
                    if 'pdf_errors' in locals():
                        if pdf_errors:
                            result_msg += f"\nPDF转换错误: {len(pdf_errors)} 个"
                    result_msg += f"\n\nWord文件保存在: {output_dir}\nPDF文件保存在: {os.path.join(output_dir, 'PDF文件')}"
                else:
                    result_msg += f"\n\n文件保存在: {output_dir}"
                
                self.root.after(0, lambda: messagebox.showinfo("完成", result_msg))
                self.root.after(0, lambda: self.log(f"批量填充任务完成: 成功{processed_count}个，错误{len(error_files)}个"))
            else:
                self.root.after(0, lambda: self.tab6_status_var.set("已停止"))
                self.root.after(0, lambda: self.log("批量填充任务已停止"))
                
        except Exception as e:
            error_msg = f"批量填充失败: {str(e)}"
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.root.after(0, lambda: self.log(f"× 批量填充任务错误: {e}"))
            
            # 记录详细的错误信息到日志
            import traceback
            error_details = traceback.format_exc()
            self.root.after(0, lambda: self.log(f"× 错误详情:\n{error_details}"))
            
            # 保存错误信息到文件
            try:
                error_file = os.path.join(output_dir if 'output_dir' in locals() else os.getcwd(), "tab6_error.log")
                with open(error_file, "w", encoding="utf-8") as f:
                    f.write(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Error: {str(e)}\n\n")
                    f.write(f"Traceback:\n{error_details}\n")
                    f.write(f"\nContext:\n")
                    f.write(f"Template: {template_file if 'template_file' in locals() else 'N/A'}\n")
                    f.write(f"Data: {data_file if 'data_file' in locals() else 'N/A'}\n")
                    f.write(f"Output: {output_dir if 'output_dir' in locals() else 'N/A'}\n")
                    f.write(f"Processed: {processed_count if 'processed_count' in locals() else 0}\n")
                self.root.after(0, lambda: self.log(f"× 错误详情已保存到: {error_file}"))
            except:
                pass
            
        finally:
            # 恢复UI状态
            self.root.after(0, self._update_tab6_ui_finish)
    
    def _update_tab6_ui_finish(self):
        """更新Tab6 UI完成状态"""
        self.tab6_is_running = False
        self.tab6_start_btn.configure(state="normal")
        self.tab6_stop_btn.configure(state="disabled", text="■ 停止任务")

    def test_vector_connection_tab5(self):
        """测试Tab5向量模型连接(Qwen)"""
        key = self.tab5_vector_api_key_var.get().strip()
        model = self.tab5_vector_model_var.get().strip() or "text-embedding-v3"
        if not key:
            messagebox.showwarning("提示", "请填写向量模型API KEY")
            return
        
        # 保存配置
        self.config["tab5_vector_api_key"] = key
        self.config["tab5_vector_model"] = model
        self.save_config()
        
        self.log(f"正在测试Qwen向量模型连接: {model} ...")
        
        try:
            response = requests.post(
                "https://dashscope.aliyuncs.com/compatible-mode/v1/embeddings",
                headers={
                    "Authorization": f"Bearer {key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": model,
                    "input": ["测试连接"],
                    "encoding_format": "float"
                },
                timeout=10
            )
            
            if response.status_code == 200:
                response_data = response.json()
                embeddings = response_data.get("data", [])
                if embeddings and embeddings[0].get("embedding"):
                    embedding_dim = len(embeddings[0]["embedding"])
                    self.log(f"✓ Qwen向量模型连接成功，返回 {embedding_dim} 维向量")
                    messagebox.showinfo("成功", f"Qwen向量模型连接成功！\n返回维度: {embedding_dim}")
                else:
                    raise ValueError(f"接口已返回200，但未获取到embedding数据: {response_data}")
            else:
                try:
                    error_data = response.json()
                    error_msg = error_data.get("message", str(error_data))
                except:
                    error_msg = f"HTTP {response.status_code}: {response.text[:200]}"
                self.log(f"× Qwen连接失败: {error_msg}")
                messagebox.showerror("连接失败", f"错误信息:\n{error_msg}")
        except Exception as e:
            self.log(f"× Qwen连接失败: {e}")
            messagebox.showerror("连接失败", f"错误信息:\n{e}")

    def stop_vector_archive_task(self):
        """停止Tab5向量归档任务"""
        if self.tab5_is_running:
            self.tab5_stop_flag = True
            self.log("🛑 正在停止向量归档任务...")
            self.tab5_stop_btn.configure(state="disabled", text="停止中...")

    def update_tab5_ui_state(self, is_running):
        """更新Tab5 UI状态"""
        self.tab5_is_running = is_running
        if is_running:
            self.tab5_start_btn.configure(state="disabled")
            self.tab5_stop_btn.configure(state="normal", text="■ 停止任务")
            self.tab5_stop_flag = False
        else:
            self.tab5_start_btn.configure(state="normal")
            self.tab5_stop_btn.configure(state="disabled", text="■ 停止任务")

    def start_vector_archive_task(self):
        """启动Tab5向量归档任务"""
        # 参数校验
        if self.tab5_df is None:
            messagebox.showwarning("配置不全", "请先加载Excel文件")
            return
        
        abs_key1 = self.tab5_abs_key1_cb.get()
        abs_key2 = self.tab5_abs_key2_cb.get()
        
        if not abs_key1 or not abs_key2:
            messagebox.showwarning("配置不全", "请选择两个绝对关键列")
            return
        
        vector_key = self.tab5_vector_api_key_var.get().strip()
        if not vector_key:
            messagebox.showwarning("配置不全", "请填写向量模型API Key")
            return
        
        source_dir = self.tab5_source_dir_var.get().strip()
        dest_dir = self.tab5_dest_dir_var.get().strip()
        
        if not source_dir or not os.path.exists(source_dir):
            messagebox.showerror("错误", "文件源路径无效")
            return
        if not dest_dir:
            messagebox.showerror("错误", "请选择归档目的地")
            return
        
        # 检查OCR配置
        ocr_key = self.tab5_ocr_api_key_var.get().strip()
        
        if not ocr_key:
            messagebox.showwarning("配置不全", "请配置Tab5独立OCR API Key")
            return
        
        # 解析页码
        page_nums = self.parse_page_numbers_tab5(self.tab5_ocr_pages_var.get())
        
        # 获取关键列配置
        key_columns = {
            'absolute': [abs_key1, abs_key2],
            'auxiliary': [
                self.tab5_aux_key1_cb.get(),
                self.tab5_aux_key2_cb.get(),
                self.tab5_aux_key3_cb.get()
            ]
        }
        
        # 获取OCR配置
        ocr_max_chars = self.tab5_ocr_max_chars_var.get().strip()
        
        self.log("--- 开始向量智能归档任务 ---")
        self.log(f"绝对关键列: {abs_key1}, {abs_key2}")
        self.log(f"辅助关键列: {[c for c in key_columns['auxiliary'] if c]}")
        self.log(f"OCR页码: {page_nums}")
        self.log("使用Tab5独立OCR配置")
        
        # 启动线程
        thread = threading.Thread(
            target=self._run_vector_archive_task,
            args=(key_columns, page_nums, ocr_max_chars),
            daemon=True
        )
        thread.start()

    def parse_page_numbers_tab5(self, page_str):
        """解析Tab5页码字符串"""
        pages = set()
        try:
            parts = page_str.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start, end = part.split('-', 1)
                    start, end = int(start.strip()), int(end.strip())
                    pages.update(range(start, end + 1))
                else:
                    pages.add(int(part.strip()))
            return sorted(list(pages))
        except:
            return [1]

    def _sanitize_folder_name_tab5(self, name):
        """清理Tab5文件夹名称中的非法字符"""
        if not name or name == 'nan':
            return "unnamed"
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            name = name.replace(char, '_')
        name = ''.join(char for char in name if ord(char) >= 32)
        name = name.strip(' .')
        if not name:
            name = "unnamed"
        return name

    def _get_file_content_with_pages_tab5(self, file_path, page_nums):
        """获取Tab5文件多页内容（PDF或图片）"""
        contents = []
        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.pdf':
                doc = fitz.open(file_path)
                total_pages = len(doc)

                for page_num in page_nums:
                    if page_num < 1 or page_num > total_pages:
                        continue
                    page = doc.load_page(page_num - 1)
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")
                    contents.append((page_num, img_bytes))

                doc.close()
            elif ext in ['.jpg', '.jpeg', '.png']:
                with open(file_path, 'rb') as f:
                    img_bytes = f.read()
                contents.append((1, img_bytes))
        except Exception as e:
            self.log(f"  Tab5文件读取错误: {e}")

        return contents

    def _ocr_local_tab5(self, img_bytes, file_path):
        """Tab5本地OCR识别"""
        try:
            if self.tab5_ocr_engine is None:
                self.log("  正在初始化Tab5本地OCR引擎(RapidOCR)，请稍候...")
                try:
                    # 动态导入 RapidOCR
                    from rapidocr import RapidOCR
                    ocr_result = [None]
                    def init_ocr():
                        try:
                            ocr_result[0] = RapidOCR()
                        except Exception as e:
                            self.log(f"  Tab5 OCR引擎初始化失败: {e}")
                            ocr_result[0] = None

                    ocr_thread = threading.Thread(target=init_ocr)
                    ocr_thread.daemon = True
                    ocr_thread.start()
                    ocr_thread.join(timeout=30)

                    if ocr_result[0] is None:
                        self.log("  ⚠️ Tab5 OCR引擎初始化超时或失败，跳过本地OCR")
                        return ""

                    self.tab5_ocr_engine = ocr_result[0]
                    self.log("  ✓ Tab5本地OCR引擎初始化完成")
                except ImportError:
                    self.log("  ⚠️ 未安装RapidOCR，跳过本地OCR")
                    self.tab5_ocr_engine = "not_installed"
                    return ""
                except Exception as e:
                    self.log(f"  ⚠️ Tab5 OCR初始化异常: {e}")
                    return ""
            
            if self.tab5_ocr_engine == "not_installed":
                self.log("  ⚠️ Tab5 OCR引擎不可用，跳过本地OCR")
                return ""

            result, _ = self.tab5_ocr_engine(img_bytes)
            if result:
                text = "".join([line[1] for line in result])
                self.log(f"  ✓ Tab5 OCR识别完成，提取 {len(text)} 字符")
                return text
            self.log("  ⚠️ Tab5 OCR未识别到文字")
            return ""
        except Exception as e:
            self.log(f"  ⚠️ Tab5本地OCR错误: {e}")
            return ""

    def _ocr_llm_fallback_tab5(self, img_bytes, ocr_max_chars="", ocr_prompt=""):
        """Tab5 LLM兜底OCR"""
        api_key = self.tab5_ocr_api_key_var.get().strip()
        if not api_key:
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

        try:
            client = OpenAI(api_key=api_key, base_url="https://ark.cn-beijing.volces.com/api/v3")
            base64_image = base64.b64encode(img_bytes).decode('utf-8')

            prompt = ocr_prompt or "请提取图片中的文字内容。直接输出识别到的文字，不要包含解释或格式标记。"
            if ocr_max_chars:
                try:
                    max_chars = int(ocr_max_chars)
                    prompt += f"\n请控制在{max_chars}个字符以内。"
                except:
                    pass

            request_params = {
                "model": self.tab5_ocr_model_var.get().strip(),
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                            }
                        ]
                    }
                ]
            }

            if ocr_max_chars:
                try:
                    request_params["max_tokens"] = int(int(ocr_max_chars) * 2)
                except:
                    pass

            response = client.chat.completions.create(**request_params)
            result_text = response.choices[0].message.content
            token_usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
            if hasattr(response, 'usage') and response.usage:
                token_usage = {
                    "prompt_tokens": getattr(response.usage, 'prompt_tokens', 0),
                    "completion_tokens": getattr(response.usage, 'completion_tokens', 0),
                    "total_tokens": getattr(response.usage, 'total_tokens', 0)
                }
            return result_text, token_usage
        except Exception as e:
            self.log(f"  Tab5 LLM OCR错误: {e}")
            return "", {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

    def _extract_pdf_metadata_tab5(self, file_path: str) -> dict:
        """提取Tab5 PDF元数据"""
        metadata = {}
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.pdf':
            try:
                doc = fitz.open(file_path)
                meta = doc.metadata
                metadata = {
                    'title': meta.get('title', ''),
                    'author': meta.get('author', ''),
                    'subject': meta.get('subject', ''),
                    'keywords': meta.get('keywords', ''),
                    'creator': meta.get('creator', ''),
                    'producer': meta.get('producer', ''),
                    'page_count': len(doc)
                }
                doc.close()
            except Exception as e:
                print(f"[Metadata] 提取PDF元数据失败: {e}")
        
        return metadata

    def _preprocess_file_tab5(self, file_path: str, page_nums: list, 
                        ocr_max_chars: str,
                        vector_manager: VectorManager, task_id: int) -> dict:
        """预处理单个文件：OCR识别 + 向量生成"""
        # 计算文件hash
        with open(file_path, 'rb') as f:
            file_hash = hashlib.sha256(f.read()).hexdigest()
        
        # 检查数据库是否已有记录
        cached = self.db.get_file_metadata(file_hash)
        if cached and not self.tab5_skip_hash_check_var.get():
            self.log(f"  [缓存命中] {os.path.basename(file_path)}")
            return cached
        
        self.log(f"  [预处理] {os.path.basename(file_path)}")
        
        # 提取PDF元数据
        pdf_metadata = self._extract_pdf_metadata_tab5(file_path)
        pdf_metadata_str = json.dumps(pdf_metadata, ensure_ascii=False) if pdf_metadata else ''
        
        contents = self._get_file_content_with_pages_tab5(file_path, page_nums)
        ocr_text = ''
        ocr_source = 'local'
        ocr_model_id = ''
        ocr_token_usage = 0
        
        for page_num, img_bytes in contents:
            text = self._ocr_local_tab5(img_bytes, file_path)
            if text:
                ocr_text += text + '\n'
                ocr_source = 'local'
            else:
                api_key = self.tab5_ocr_api_key_var.get().strip()
                model_id = self.tab5_ocr_model_var.get()
                
                if api_key:
                    llm_text, token_usage = self._ocr_llm_fallback_tab5(
                        img_bytes,
                        ocr_max_chars,
                        self.tab5_ocr_prompt_var.get().strip()
                    )
                    if llm_text:
                        ocr_text += llm_text + '\n'
                        ocr_source = 'llm'
                        ocr_model_id = model_id
                        ocr_token_usage += token_usage.get('total_tokens', 0)
        
        # 获取文件名（不含扩展名）
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        file_type = os.path.splitext(file_path)[1].lower()
        
        # 向量化
        content_vec = None
        name_vec = None
        meta_vec = None
        vector_token_usage = 0
        
        texts_to_embed = []
        if ocr_text:
            texts_to_embed.append(ocr_text[:1000])  # 限制长度
        texts_to_embed.append(file_name)
        if pdf_metadata_str:
            texts_to_embed.append(pdf_metadata_str[:500])
        
        if texts_to_embed:
            vectors, tokens = vector_manager.get_embedding(texts_to_embed)
            vector_token_usage = tokens
            
            if len(vectors) >= 1:
                content_vec = VectorManager.vector_to_bytes(vectors[0])
            if len(vectors) >= 2:
                name_vec = VectorManager.vector_to_bytes(vectors[1])
            if len(vectors) >= 3:
                meta_vec = VectorManager.vector_to_bytes(vectors[2])
        
        # 保存到数据库
        self.db.save_file_metadata(
            file_hash, file_path, file_name, file_type, pdf_metadata_str,
            ocr_text, ocr_source, ocr_model_id, ocr_token_usage,
            content_vec, name_vec, meta_vec, vector_token_usage
        )
        
        return {
            'file_hash': file_hash,
            'file_path': file_path,
            'file_name': file_name,
            'file_type': file_type,
            'pdf_metadata': pdf_metadata_str,
            'ocr_text': ocr_text,
            'ocr_source': ocr_source,
            'ocr_model_id': ocr_model_id,
            'ocr_token_usage': ocr_token_usage,
            'content_vector': content_vec,
            'name_vector': name_vec,
            'meta_vector': meta_vec,
            'vector_token_usage': vector_token_usage,
            'is_classified': False
        }

    def _run_vector_archive_task(self, key_columns: dict, page_nums: list, 
                                   ocr_max_chars: str):
        """执行向量归档任务主循环"""
        try:
            self.update_tab5_ui_state(True)
            
            # 初始化向量管理器
            vector_api_key = self.tab5_vector_api_key_var.get().strip()
            vector_model = self.tab5_vector_model_var.get().strip() or "text-embedding-v3"
            vector_manager = VectorManager(vector_api_key, vector_model)
            
            # 记录任务开始
            task_id = self.db.log_task_start('vector_archive')
            
            source_dir = self.tab5_source_dir_var.get().strip()
            dest_dir = self.tab5_dest_dir_var.get().strip()
            
            # 遍历源目录获取所有文件
            all_files = []
            for root, dirs, files in os.walk(source_dir):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in ['.pdf', '.jpg', '.jpeg', '.png']:
                        all_files.append(file_path)
            
            self.log(f"找到 {len(all_files)} 个待处理文件")
            
            # 阶段1：预处理所有文件（OCR + 向量化）
            self.log("--- 阶段1: 文件预处理和向量化 ---")
            file_metadata_list = []
            for i, file_path in enumerate(all_files):
                if self.tab5_stop_flag:
                    break
                
                try:
                    metadata = self._preprocess_file_tab5(
                        file_path, page_nums, ocr_max_chars,
                        vector_manager, task_id
                    )
                    file_metadata_list.append(metadata)
                    
                    if (i + 1) % 10 == 0:
                        self.log(f"  已处理 {i+1}/{len(all_files)} 个文件")
                except Exception as e:
                    self.log(f"  × 预处理失败 {os.path.basename(file_path)}: {e}")
            
            self.log(f"--- 预处理完成: {len(file_metadata_list)} 个文件 ---")
            
            # 阶段2：反向匹配（文件→数据行）
            # 设计：让每个文件寻找最匹配的数据行(argmax)，天然支持"一条数据对应N个文件"
            self.log("--- 阶段2: 反向匹配（文件→数据行） ---")
            
            # 权重配置
            base_weights = {
                key_columns['absolute'][0]: 100,
                key_columns['absolute'][1]: 80,
            }
            for i, aux_col in enumerate(key_columns['auxiliary']):
                if aux_col:
                    weights = [50, 30, 15]
                    base_weights[aux_col] = weights[i]
            
            source_coefficients = {
                'content': 1.0,
                'name': 1.5,
                'meta': 0.7
            }
            
            # 全局底线阈值：最佳匹配的最高余弦相似度低于此值则视为未匹配
            BASE_THRESHOLD = 0.50
            
            # 2a: 预计算所有数据行的关键列向量（一次性embedding，避免重复API调用）
            self.log("  正在预计算数据行向量...")
            all_key_cols = [c for c in key_columns['absolute'] + key_columns['auxiliary'] if c]
            row_data_list = []
            
            for idx, row in self.tab5_df.iterrows():
                if self.tab5_stop_flag:
                    break
                
                key_values = {}
                for col in all_key_cols:
                    if col in row:
                        val = str(row[col]) if pd.notna(row[col]) else ''
                        if val:
                            key_values[col] = val
                
                if not key_values:
                    continue
                
                texts = list(key_values.values())
                cols = list(key_values.keys())
                vecs, _ = vector_manager.get_embedding(texts)
                
                key_vectors = {}
                for ci, col_name in enumerate(cols):
                    if ci < len(vecs):
                        key_vectors[col_name] = vecs[ci]
                
                if key_vectors:
                    row_data_list.append({
                        'row_idx': idx,
                        'key_values': key_values,
                        'key_vectors': key_vectors
                    })
                
                if (idx + 1) % 50 == 0:
                    self.log(f"  已预计算 {idx+1} 行...")
            
            self.log(f"  预计算完成: {len(row_data_list)} 条有效数据行")
            
            # 2b: 反向匹配 - 遍历每个文件，找其最佳数据行 (argmax + 底线阈值)
            self.log(f"  反向匹配中（底线阈值: {BASE_THRESHOLD}）...")
            
            categorized = {}
            unmatched_files = []
            
            for fi, file_meta in enumerate(file_metadata_list):
                if self.tab5_stop_flag:
                    break
                
                content_vec = file_meta.get('content_vector')
                name_vec = file_meta.get('name_vector')
                meta_vec = file_meta.get('meta_vector')
                
                if isinstance(content_vec, bytes):
                    content_vec = vector_manager.bytes_to_vector(content_vec)
                if isinstance(name_vec, bytes):
                    name_vec = vector_manager.bytes_to_vector(name_vec)
                if isinstance(meta_vec, bytes):
                    meta_vec = vector_manager.bytes_to_vector(meta_vec)
                
                file_sources = []
                if content_vec is not None:
                    file_sources.append(('content', content_vec))
                if name_vec is not None:
                    file_sources.append(('name', name_vec))
                if meta_vec is not None:
                    file_sources.append(('meta', meta_vec))
                
                if not file_sources:
                    unmatched_files.append(file_meta)
                    continue
                
                best_total_score = -1
                best_max_sim = 0
                best_row_data = None
                best_matched_cols = []
                best_details = []
                
                for row_data in row_data_list:
                    total_score = 0
                    max_sim = 0
                    matched_cols = []
                    details = []
                    
                    for col_name, key_vec in row_data['key_vectors'].items():
                        bw = base_weights.get(col_name, 50)
                        col_best_sim = 0
                        col_best_source = None
                        
                        for source_name, file_vec in file_sources:
                            sim = vector_manager.cosine_similarity(key_vec, file_vec)
                            if sim > col_best_sim:
                                col_best_sim = sim
                                col_best_source = source_name
                        
                        if col_best_sim > 0 and col_best_source:
                            score = bw * source_coefficients[col_best_source] * col_best_sim
                            total_score += score
                            if col_best_sim > max_sim:
                                max_sim = col_best_sim
                            matched_cols.append(col_name)
                            details.append({
                                'column': col_name,
                                'value': row_data['key_values'][col_name],
                                'source': col_best_source,
                                'similarity': round(col_best_sim, 4),
                                'score': round(score, 2)
                            })
                    
                    if total_score > best_total_score:
                        best_total_score = total_score
                        best_max_sim = max_sim
                        best_row_data = row_data
                        best_matched_cols = matched_cols
                        best_details = details
                
                abs_matched = set(best_matched_cols) & set(key_columns['absolute'])
                if abs_matched and best_max_sim >= BASE_THRESHOLD and best_row_data:
                    row_idx = best_row_data['row_idx']
                    if row_idx not in categorized:
                        categorized[row_idx] = []
                    categorized[row_idx].append({
                        'file_meta': file_meta,
                        'score': best_total_score,
                        'max_similarity': best_max_sim,
                        'matched_columns': best_matched_cols,
                        'weight_details': json.dumps(best_details, ensure_ascii=False),
                        'key_values': best_row_data['key_values']
                    })
                else:
                    unmatched_files.append(file_meta)
                
                if (fi + 1) % 10 == 0:
                    self.log(f"  已匹配 {fi+1}/{len(file_metadata_list)} 个文件")
            
            matched_total = sum(len(v) for v in categorized.values())
            self.log(f"--- 反向匹配完成: {matched_total} 个文件→{len(categorized)} 条数据行, {len(unmatched_files)} 个未匹配 ---")
            
            # 阶段3：归档
            self.log("--- 阶段3: 归档文件 ---")
            
            processed_count = 0
            unclassified_dir = os.path.join(dest_dir, "0未匹配待处理")
            
            for row_idx, matched_list in categorized.items():
                if self.tab5_stop_flag:
                    break
                
                row = self.tab5_df.iloc[row_idx]
                
                # 使用Tab5独立的文件夹命名配置
                folder_cols = [
                    self.tab5_folder_name_col1_cb.get(),
                    self.tab5_folder_name_col2_cb.get(),
                    self.tab5_folder_name_col3_cb.get()
                ]
                
                folder_parts = []
                for col in folder_cols:
                    if col and col in row and pd.notna(row[col]):
                        folder_parts.append(str(row[col]))
                
                if folder_parts:
                    folder_name = '_'.join(folder_parts)
                else:
                    # 如果没有配置文件夹命名，使用绝对关键列作为后备
                    abs_key1 = key_columns['absolute'][0]
                    abs_key2 = key_columns['absolute'][1]
                    if abs_key1 and abs_key1 in row and pd.notna(row[abs_key1]):
                        folder_name = str(row[abs_key1])
                    elif abs_key2 and abs_key2 in row and pd.notna(row[abs_key2]):
                        folder_name = str(row[abs_key2])
                    else:
                        folder_name = f"未知_{row_idx}"
                
                folder_name = self._sanitize_folder_name_tab5(folder_name)
                target_folder = os.path.join(dest_dir, folder_name)
                
                matched_list.sort(key=lambda x: x['score'], reverse=True)
                
                for match_info in matched_list:
                    if self.tab5_stop_flag:
                        break
                    
                    fm = match_info['file_meta']
                    file_hash = fm['file_hash']
                    
                    # 使用Tab5独立的文件命名配置
                    file_cols = [
                        self.tab5_file_name_col1_cb.get(),
                        self.tab5_file_name_col2_cb.get(),
                        self.tab5_file_name_col3_cb.get()
                    ]
                    
                    file_parts = []
                    for col in file_cols:
                        if col and col in row and pd.notna(row[col]):
                            file_parts.append(str(row[col]))
                    
                    if file_parts:
                        # 使用配置的列组合作为文件名
                        base_name = '_'.join(file_parts)
                        # 保留原扩展名
                        ext = fm['file_type']
                        target_filename = base_name + ext
                    else:
                        # 如果没有配置文件命名，使用原文件名
                        target_filename = fm['file_name'] + fm['file_type']
                    
                    self.db.save_classification_result(
                        task_id, file_hash, row_idx,
                        match_info['score'],
                        json.dumps(match_info['matched_columns'], ensure_ascii=False),
                        match_info['weight_details'],
                        False,
                        target_folder,
                        target_filename
                    )
                    
                    try:
                        os.makedirs(target_folder, exist_ok=True)
                        src_path = fm['file_path']
                        dst_path = os.path.join(target_folder, target_filename)
                        
                        counter = 1
                        base_name = os.path.splitext(target_filename)[0]
                        ext = fm['file_type']
                        while os.path.exists(dst_path):
                            target_filename = f"{base_name}_{counter}{ext}"
                            dst_path = os.path.join(target_folder, target_filename)
                            counter += 1
                        
                        shutil.move(src_path, dst_path)
                        self.db.update_file_classified_status(file_hash, True)
                        
                        processed_count += 1
                        self.log(f"✓ 归档: {os.path.basename(src_path)} -> {folder_name}/ (相似度: {match_info['max_similarity']:.4f})")
                    except Exception as e:
                        self.log(f"× 归档失败 {os.path.basename(fm['file_path'])}: {e}")
            
            # 处理未匹配文件（低于底线阈值，移入待处理目录）
            unclassified_count = 0
            for file_meta in unmatched_files:
                if self.tab5_stop_flag:
                    break
                try:
                    os.makedirs(unclassified_dir, exist_ok=True)
                    src_path = file_meta['file_path']
                    target_filename = file_meta['file_name'] + file_meta['file_type']
                    dst_path = os.path.join(unclassified_dir, target_filename)
                    
                    counter = 1
                    base_name = os.path.splitext(target_filename)[0]
                    ext = file_meta['file_type']
                    while os.path.exists(dst_path):
                        target_filename = f"{base_name}_{counter}{ext}"
                        dst_path = os.path.join(unclassified_dir, target_filename)
                        counter += 1
                    
                    shutil.move(src_path, dst_path)
                    unclassified_count += 1
                except Exception as e:
                    self.log(f"× 移动未匹配文件失败: {e}")
            
            # 统计Token使用情况
            token_stats = self.db.get_task_token_usage(task_id)
            
            # 更新任务结束
            self.db.log_task_end(task_id, 'completed' if not self.tab5_stop_flag else 'stopped',
                                processed_count, processed_count, token_stats['total_tokens'])
            
            # 打印统计信息
            self.log("--- 向量归档任务完成 ---")
            self.log(f"📊 处理文件: {processed_count}")
            self.log(f"📊 匹配数据行: {len(categorized)}")
            self.log(f"📊 未匹配文件: {unclassified_count}")
            self.log(f"📊 OCR Token: {token_stats['ocr_tokens']}")
            self.log(f"📊 向量 Token: {token_stats['vector_tokens']}")
            
            if not self.tab5_stop_flag:
                self.root.after(0, lambda: messagebox.showinfo(
                    "完成",
                    f"向量归档任务完成！\n\n"
                    f"已归档: {processed_count} 个文件\n"
                    f"匹配数据行: {len(categorized)} 条\n"
                    f"未匹配: {unclassified_count} 个\n\n"
                    f"Token使用:\n"
                    f"  OCR: {token_stats['ocr_tokens']}\n"
                    f"  向量: {token_stats['vector_tokens']}"
                ))
            
        except Exception as e:
            self.log(f"向量归档任务错误: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("错误", f"任务失败:\n{e}"))
        finally:
            self.update_tab5_ui_state(False)



if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = FileToolApp(root)
        root.mainloop()
    except Exception as e:
        # 简单的错误处理，避免递归
        import traceback
        error_details = traceback.format_exc()
        try:
            with open("crash_log.txt", "w", encoding="utf-8") as f:
                f.write(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Error: {e}\n\n")
                f.write(f"Traceback:\n{error_details}\n")
        except:
            pass
        print(f"程序崩溃: {e}")
        print(f"详细信息已保存到 crash_log.txt")
