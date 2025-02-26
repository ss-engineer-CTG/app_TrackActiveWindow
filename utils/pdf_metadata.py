# pdf_metadata.py
import os
from typing import Dict, Any, Optional
from PyPDF2 import PdfReader

class PDFMetadataExtractor:
    """PDFファイルからメタデータを抽出するユーティリティクラス"""
    
    def __init__(self):
        # メタデータのキャッシュ
        self._metadata_cache: Dict[str, Dict[str, Any]] = {}
    
    def extract_metadata(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """PDFファイルからメタデータを抽出"""
        # キャッシュにあれば返す
        if pdf_path in self._metadata_cache:
            return self._metadata_cache[pdf_path]
        
        # ファイルが存在しない場合はNoneを返す
        if not os.path.exists(pdf_path) or not pdf_path.lower().endswith('.pdf'):
            return None
        
        try:
            # PDFファイルを開く
            reader = PdfReader(pdf_path)
            
            # 基本情報を取得
            info = reader.metadata
            if info:
                # 一般的なメタデータフィールドを抽出
                metadata = {
                    'title': info.get('/Title', ''),
                    'author': info.get('/Author', ''),
                    'subject': info.get('/Subject', ''),
                    'creator': info.get('/Creator', ''),
                    'producer': info.get('/Producer', ''),
                    'creation_date': info.get('/CreationDate', ''),
                    'modification_date': info.get('/ModDate', ''),
                    'page_count': len(reader.pages)
                }
                
                # キャッシュに保存
                self._metadata_cache[pdf_path] = metadata
                return metadata
        except Exception as e:
            print(f"Error extracting PDF metadata from {pdf_path}: {e}")
        
        return None
    
    def clear_cache(self):
        """メタデータキャッシュをクリア"""
        self._metadata_cache.clear()
        
    def get_cached_paths(self):
        """キャッシュされているパスのリストを取得"""
        return list(self._metadata_cache.keys())
        
    def cache_size(self):
        """キャッシュサイズを取得"""
        return len(self._metadata_cache)