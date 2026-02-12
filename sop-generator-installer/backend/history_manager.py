import json
import os
from datetime import datetime
from typing import List, Dict, Optional

class HistoryManager:
    """Manages history of previously entered SOP metadata per user"""

    def __init__(self, history_dir: str = 'history'):
        self.history_dir = history_dir
        self.current_user = None
        self.history = []

        # Create history directory if it doesn't exist
        os.makedirs(history_dir, exist_ok=True)

    def set_user(self, user_id: str):
        """Set the current user and load their history"""
        self.current_user = user_id
        self.history = self._load_history()

    def _get_history_file(self) -> str:
        """Get the history file path for current user"""
        if self.current_user:
            return os.path.join(self.history_dir, f'history_{self.current_user}.json')
        return os.path.join(self.history_dir, 'history_default.json')

    def _load_history(self) -> List[Dict]:
        """Load history from JSON file"""
        history_file = self._get_history_file()
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                return []
        return []

    def _save_history(self):
        """Save history to JSON file"""
        history_file = self._get_history_file()
        try:
            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, indent=2, ensure_ascii=False)
        except IOError as e:
            print(f"Warning: Could not save history: {e}")

    def add_entry(self, metadata: Dict):
        """
        Add a new entry to history

        Args:
            metadata: Dictionary containing:
                - process_name
                - process_code
                - purpose
                - scope
                - abbreviations
                - referenced_docs
        """
        # Create entry with timestamp
        entry = {
            'timestamp': datetime.now().isoformat(),
            'process_name': metadata.get('process_name', ''),
            'process_code': metadata.get('process_code', ''),
            'purpose': metadata.get('purpose', ''),
            'scope': metadata.get('scope', ''),
            'abbreviations_list': metadata.get('abbreviations_list', []),
            'references_list': metadata.get('references_list', []),
            'general_policies_list': metadata.get('general_policies_list', [])
        }

        # Check if this exact entry already exists (excluding timestamp)
        entry_without_timestamp = {k: v for k, v in entry.items() if k != 'timestamp'}
        for existing in self.history:
            existing_without_timestamp = {k: v for k, v in existing.items() if k != 'timestamp'}
            if existing_without_timestamp == entry_without_timestamp:
                # Update timestamp of existing entry
                existing['timestamp'] = entry['timestamp']
                self._save_history()
                return

        # Add new entry at the beginning (most recent first)
        self.history.insert(0, entry)

        # Keep only last 50 entries
        self.history = self.history[:50]

        self._save_history()

    def get_all(self) -> List[Dict]:
        """Get all history entries (most recent first)"""
        return self.history

    def get_entry(self, index: int) -> Optional[Dict]:
        """Get a specific history entry by index"""
        if 0 <= index < len(self.history):
            return self.history[index]
        return None

    def search(self, query: str) -> List[Dict]:
        """Search history by process name or code"""
        query_lower = query.lower()
        results = []
        for entry in self.history:
            if (query_lower in entry.get('process_name', '').lower() or
                query_lower in entry.get('process_code', '').lower()):
                results.append(entry)
        return results

    def clear(self):
        """Clear all history for current user"""
        self.history = []
        self._save_history()
