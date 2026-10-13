import sqlite3
import json
import logging
from typing import Dict, List, Any, Tuple, Optional
from datetime import datetime
import os

logger = logging.getLogger("db_manager")

class DatabaseManager:
    """Manages SQLite local database for storing profile history."""

    def __init__(self, db_path: str = "app_history.db"):
        self.db_path = os.path.abspath(db_path)
        self.init_db()

    def init_db(self):
        """Creates the necessary tables if they do not exist."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS history_profiles (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        timestamp TEXT NOT NULL,
                        nama_bu TEXT NOT NULL,
                        data_json TEXT NOT NULL
                    )
                ''')
                conn.commit()
        except Exception as e:
            logger.error(f"Failed to initialize database: {e}")

    def save_profile(self, nama_bu: str, profile_data: Dict[str, Any]) -> Tuple[bool, str]:
        """Saves a profile (variables and images) into the history database."""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data_json = json.dumps(profile_data, ensure_ascii=False)
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO history_profiles (timestamp, nama_bu, data_json)
                    VALUES (?, ?, ?)
                ''', (timestamp, nama_bu, data_json))
                conn.commit()
            return True, "Profile saved to history successfully."
        except Exception as e:
            logger.error(f"Failed to save profile to db: {e}")
            return False, str(e)

    def get_all_profiles(self) -> List[Dict[str, Any]]:
        """Retrieves all profiles ordered by most recent first."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('SELECT * FROM history_profiles ORDER BY id DESC')
                rows = cursor.fetchall()
                
                results = []
                for row in rows:
                    results.append({
                        "id": row["id"],
                        "timestamp": row["timestamp"],
                        "nama_bu": row["nama_bu"],
                        "data_json": json.loads(row["data_json"])
                    })
                return results
        except Exception as e:
            logger.error(f"Failed to fetch profiles: {e}")
            return []

    def search_profiles(self, keyword: str) -> List[Dict[str, Any]]:
        """Searches profiles by nama_bu or timestamp."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                search_term = f"%{keyword}%"
                cursor.execute('''
                    SELECT * FROM history_profiles 
                    WHERE nama_bu LIKE ? OR timestamp LIKE ?
                    ORDER BY id DESC
                ''', (search_term, search_term))
                rows = cursor.fetchall()
                
                results = []
                for row in rows:
                    results.append({
                        "id": row["id"],
                        "timestamp": row["timestamp"],
                        "nama_bu": row["nama_bu"],
                        "data_json": json.loads(row["data_json"])
                    })
                return results
        except Exception as e:
            logger.error(f"Failed to search profiles: {e}")
            return []

    def delete_profile(self, profile_id: int) -> bool:
        """Deletes a specific profile from history."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM history_profiles WHERE id = ?', (profile_id,))
                conn.commit()
            return True
        except Exception as e:
            logger.error(f"Failed to delete profile: {e}")
            return False
