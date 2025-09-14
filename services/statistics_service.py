import json
import os
from datetime import datetime
from threading import Lock

class StatisticsService:
    def __init__(self, stats_file='statistics.json'):
        self.stats_file = stats_file
        self.lock = Lock()
        self.stats = self._load_stats()
    
    def _load_stats(self):
        """Load statistics from file or create initial structure"""
        if os.path.exists(self.stats_file):
            try:
                with open(self.stats_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        
        # Return initial structure
        return {
            "export_stats": {
                "excel": 0,
                "word": 0,
                "image": 0,
                "print": 0
            },
            "import_stats": {
                "total": 0
            },
            "last_updated": None
        }
    
    def _save_stats(self):
        """Save statistics to file"""
        self.stats["last_updated"] = datetime.now().isoformat()
        with self.lock:
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(self.stats, f, ensure_ascii=False, indent=2)
    
    def track_export(self, export_type):
        """Track export actions"""
        if export_type in self.stats["export_stats"]:
            self.stats["export_stats"][export_type] += 1
            self._save_stats()
    
    def track_import(self):
        """Track import actions"""
        self.stats["import_stats"]["total"] += 1
        self._save_stats()
    
    def get_stats(self):
        """Get current statistics"""
        # Calculate total usage
        export_total = sum(self.stats["export_stats"].values())
        import_total = self.stats["import_stats"]["total"]
        self.stats["total_usage"] = export_total + import_total
        return self.stats

# Create a global instance
statistics_service = StatisticsService()