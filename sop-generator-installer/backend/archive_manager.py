"""
Archive Manager for SOP Generator
Handles storage and retrieval of BPMN files and generated Word documents per user
"""

import sqlite3
import os
import shutil
from datetime import datetime
from typing import List, Dict, Optional
import json

class ArchiveManager:
    def __init__(self, archive_dir: str = "archives", db_path: str = "archive.db"):
        """
        Initialize archive manager

        Args:
            archive_dir: Base directory for storing archived files
            db_path: Path to SQLite database
        """
        self.archive_dir = archive_dir
        self.db_path = db_path

        # Create archive directory if it doesn't exist
        os.makedirs(archive_dir, exist_ok=True)

        # Initialize database
        self._init_database()

    def _init_database(self):
        """Create database tables if they don't exist"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Create archives table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS archives (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                process_name TEXT NOT NULL,
                bpmn_filename TEXT NOT NULL,
                docx_filename TEXT NOT NULL,
                bpmn_path TEXT NOT NULL,
                docx_path TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        ''')

        # Create index on user_id for faster queries
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_user_id ON archives(user_id)
        ''')

        conn.commit()
        conn.close()

    def save_archive(self, user_id: str, process_name: str, bpmn_file_path: str, docx_file_path: str) -> int:
        """
        Save BPMN and Word files to user's archive

        Args:
            user_id: User identifier
            process_name: Name of the process
            bpmn_file_path: Path to source BPMN file
            docx_file_path: Path to source Word file

        Returns:
            Archive ID
        """
        # Create timestamp-based folder name
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        # Create user directory if it doesn't exist
        user_dir = os.path.join(self.archive_dir, user_id)
        os.makedirs(user_dir, exist_ok=True)

        # Create archive folder
        archive_folder = os.path.join(user_dir, timestamp)
        os.makedirs(archive_folder, exist_ok=True)

        # Get file names
        bpmn_filename = os.path.basename(bpmn_file_path)
        docx_filename = os.path.basename(docx_file_path)

        # Copy files to archive folder
        bpmn_dest = os.path.join(archive_folder, bpmn_filename)
        docx_dest = os.path.join(archive_folder, docx_filename)

        shutil.copy2(bpmn_file_path, bpmn_dest)
        shutil.copy2(docx_file_path, docx_dest)

        # Save to database
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO archives (user_id, process_name, bpmn_filename, docx_filename,
                                  bpmn_path, docx_path, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            user_id,
            process_name,
            bpmn_filename,
            docx_filename,
            bpmn_dest,
            docx_dest,
            datetime.now().isoformat()
        ))

        archive_id = cursor.lastrowid
        conn.commit()
        conn.close()

        return archive_id

    def get_user_archives(self, user_id: str, limit: int = 100) -> List[Dict]:
        """
        Get all archives for a user

        Args:
            user_id: User identifier
            limit: Maximum number of archives to return

        Returns:
            List of archive dictionaries
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('''
            SELECT id, process_name, bpmn_filename, docx_filename, created_at
            FROM archives
            WHERE user_id = ?
            ORDER BY created_at DESC
            LIMIT ?
        ''', (user_id, limit))

        archives = []
        for row in cursor.fetchall():
            archives.append({
                'id': row['id'],
                'process_name': row['process_name'],
                'bpmn_filename': row['bpmn_filename'],
                'docx_filename': row['docx_filename'],
                'created_at': row['created_at']
            })

        conn.close()
        return archives

    def get_archive(self, archive_id: int) -> Optional[Dict]:
        """
        Get a specific archive by ID

        Args:
            archive_id: Archive ID

        Returns:
            Archive dictionary or None if not found
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('''
            SELECT id, user_id, process_name, bpmn_filename, docx_filename,
                   bpmn_path, docx_path, created_at
            FROM archives
            WHERE id = ?
        ''', (archive_id,))

        row = cursor.fetchone()
        conn.close()

        if row:
            return {
                'id': row['id'],
                'user_id': row['user_id'],
                'process_name': row['process_name'],
                'bpmn_filename': row['bpmn_filename'],
                'docx_filename': row['docx_filename'],
                'bpmn_path': row['bpmn_path'],
                'docx_path': row['docx_path'],
                'created_at': row['created_at']
            }
        return None

    def delete_archive(self, archive_id: int, user_id: str) -> bool:
        """
        Delete an archive (only if it belongs to the user)

        Args:
            archive_id: Archive ID
            user_id: User ID (for security check)

        Returns:
            True if deleted, False if not found or doesn't belong to user
        """
        archive = self.get_archive(archive_id)

        if not archive or archive['user_id'] != user_id:
            return False

        # Delete files
        try:
            if os.path.exists(archive['bpmn_path']):
                os.remove(archive['bpmn_path'])
            if os.path.exists(archive['docx_path']):
                os.remove(archive['docx_path'])

            # Try to delete the folder if empty
            folder = os.path.dirname(archive['bpmn_path'])
            if os.path.exists(folder) and not os.listdir(folder):
                os.rmdir(folder)
        except Exception as e:
            print(f"Error deleting files: {e}")

        # Delete from database
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM archives WHERE id = ?', (archive_id,))

        conn.commit()
        conn.close()

        return True

    def get_file_path(self, archive_id: int, file_type: str) -> Optional[str]:
        """
        Get file path for an archive

        Args:
            archive_id: Archive ID
            file_type: 'bpmn' or 'docx'

        Returns:
            File path or None if not found
        """
        archive = self.get_archive(archive_id)

        if not archive:
            return None

        if file_type == 'bpmn':
            return archive['bpmn_path'] if os.path.exists(archive['bpmn_path']) else None
        elif file_type == 'docx':
            return archive['docx_path'] if os.path.exists(archive['docx_path']) else None

        return None
