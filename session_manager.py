"""
session_manager.py
==================
Handles session state and logging.

FIXED:
- Added current_script field to SessionState
- Added PowerPoint logging methods
- Fixed type hints

Version: 1.1.0
"""

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional
from datetime import datetime
from enum import Enum
import json


class LogLevel(Enum):
    """Log levels for session logging."""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    SUCCESS = "SUCCESS"


@dataclass
class LogEntry:
    """A single log entry."""
    timestamp: datetime
    level: LogLevel
    message: str
    details: Optional[Dict[str, Any]] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            "timestamp": self.timestamp.isoformat(),
            "level": self.level.value,
            "message": self.message,
            "details": self.details
        }
    
    def __str__(self) -> str:
        level_icons = {
            LogLevel.DEBUG: "🔍",
            LogLevel.INFO: "ℹ️",
            LogLevel.WARNING: "⚠️",
            LogLevel.ERROR: "❌",
            LogLevel.SUCCESS: "✓"
        }
        icon = level_icons.get(self.level, "•")
        return f"[{self.timestamp.strftime('%H:%M:%S')}] {icon} {self.message}"


@dataclass
class SessionState:
    """
    Holds the current session state.
    
    FIXED: Added current_script field for PowerPoint scripts
    """
    session_id: str = ""
    user_id: Optional[str] = None
    start_time: datetime = field(default_factory=datetime.now)
    
    # File tracking
    current_file: Optional[str] = None
    current_workbook: Optional[Any] = None
    current_presentation: Optional[Any] = None  # ADDED: For PowerPoint
    current_script: Optional[str] = None  # ADDED: For script storage
    
    # History
    files_processed: List[str] = field(default_factory=list)
    operations_performed: List[str] = field(default_factory=list)
    
    # Metadata
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def reset(self):
        """Reset state but keep session info."""
        self.current_file = None
        self.current_workbook = None
        self.current_presentation = None
        self.current_script = None
        self.metadata = {}


class SessionManager:
    """
    Manages session state and provides logging functionality.
    """
    
    def __init__(self, session_id: Optional[str] = None):
        """Initialize session manager."""
        self.state = SessionState(
            session_id=session_id or self._generate_session_id()
        )
        self.logs: List[LogEntry] = []
        self._max_logs = 1000
    
    def _generate_session_id(self) -> str:
        """Generate unique session ID."""
        return f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    def _add_log(self, level: LogLevel, message: str, details: Optional[Dict] = None):
        """Add a log entry."""
        entry = LogEntry(
            timestamp=datetime.now(),
            level=level,
            message=message,
            details=details
        )
        self.logs.append(entry)
        
        # Trim old logs
        if len(self.logs) > self._max_logs:
            self.logs = self.logs[-self._max_logs:]
        
        # Print to console
        print(str(entry))
    
    # -------------------------------------------------------------------------
    # LOGGING METHODS
    # -------------------------------------------------------------------------
    
    def debug(self, message: str, details: Optional[Dict] = None):
        """Log debug message."""
        self._add_log(LogLevel.DEBUG, message, details)
    
    def info(self, message: str, details: Optional[Dict] = None):
        """Log info message."""
        self._add_log(LogLevel.INFO, message, details)
    
    def warning(self, message: str, details: Optional[Dict] = None):
        """Log warning message."""
        self._add_log(LogLevel.WARNING, message, details)
    
    def error(self, message: str, details: Optional[Dict] = None):
        """Log error message."""
        self._add_log(LogLevel.ERROR, message, details)
    
    def success(self, message: str, details: Optional[Dict] = None):
        """Log success message."""
        self._add_log(LogLevel.SUCCESS, message, details)
    

    # Add to SessionManager class in session_manager.py

    def set_prompt(self, prompt: str):
        """Store the current generated prompt."""
        if prompt is None:
            self.warning("Attempted to store None prompt - check prompt generation")
            prompt = ""
        self.set_metadata("current_prompt", prompt)
        self.info(f"Prompt stored ({len(prompt)} characters)")


    def set_innovation_context(self, context: Dict[str, Any]):
        """Store the innovation context."""
        self.set_metadata("innovation_context", context)
        innovation_name = context.get("innovation_name", "Unknown")
        self.info(f"Innovation context set: {innovation_name}")

    def set_script(self, script: str):
        """Store the current script (generic - Excel/Word/PowerPoint)."""
        self.state.current_script = script
        self.info(f"Script loaded ({len(script)} characters)")

    def set_error(self, error_details: str):
        """Log and store error details."""
        self.set_metadata("last_error", error_details)
        self.error(f"Error recorded: {error_details[:100]}...")

    def log_file_saved(self, filepath: str):
        """Log that a file was successfully saved."""
        self.success(f"File saved: {filepath}")
        self.state.files_processed.append(filepath)
        self.state.operations_performed.append(f"Saved: {filepath}")



    # -------------------------------------------------------------------------
    # FILE TRACKING
    # -------------------------------------------------------------------------
    
    def set_current_file(self, filepath: str):
        """Set the current file being processed."""
        self.state.current_file = filepath
        if filepath not in self.state.files_processed:
            self.state.files_processed.append(filepath)
        self.info(f"Working with file: {filepath}")
    
    def set_current_workbook(self, workbook: Any):
        """Set the current Excel workbook."""
        self.state.current_workbook = workbook
        self.debug("Workbook loaded into session")
    
    def set_current_presentation(self, presentation: Any):
        """Set the current PowerPoint presentation."""
        self.state.current_presentation = presentation
        self.debug("Presentation loaded into session")
    
    # -------------------------------------------------------------------------
    # POWERPOINT SPECIFIC METHODS (ADDED)
    # -------------------------------------------------------------------------
    
    def set_pptx_script(self, script: str):
        """Store the current PowerPoint script."""
        self.state.current_script = script
        self.info(f"PowerPoint script loaded ({len(script)} characters)")
    
    def log_pptx_generated(self, slide_count: int, filepath: str):
        """Log PowerPoint generation success."""
        self.success(f"PowerPoint generated: {slide_count} slides → {filepath}")
        self.state.operations_performed.append(f"Generated PPTX: {filepath}")
    
    def log_pptx_error(self, error: str):
        """Log PowerPoint generation error."""
        self.error(f"PowerPoint error: {error}")
    
    # -------------------------------------------------------------------------
    # EXCEL SPECIFIC METHODS
    # -------------------------------------------------------------------------
    
    def log_excel_operation(self, operation: str, details: Optional[Dict] = None):
        """Log an Excel operation."""
        self.info(f"Excel: {operation}", details)
        self.state.operations_performed.append(f"Excel: {operation}")
    
    def log_excel_error(self, error: str):
        """Log Excel error."""
        self.error(f"Excel error: {error}")
    
    # -------------------------------------------------------------------------
    # IMPORT OPERATIONS
    # -------------------------------------------------------------------------
    
    def log_import(self, source: str, record_count: int):
        """Log data import."""
        self.success(f"Imported {record_count} records from {source}")
        self.state.operations_performed.append(f"Import: {source} ({record_count} records)")
    
    def log_import_error(self, source: str, error: str):
        """Log import error."""
        self.error(f"Import failed from {source}: {error}")
    
    # -------------------------------------------------------------------------
    # METADATA
    # -------------------------------------------------------------------------
    
    def set_metadata(self, key: str, value: Any):
        """Set session metadata."""
        self.state.metadata[key] = value
    
    def get_metadata(self, key: str, default: Any = None) -> Any:
        """Get session metadata."""
        return self.state.metadata.get(key, default)
    
    # -------------------------------------------------------------------------
    # LOG ACCESS
    # -------------------------------------------------------------------------
    
    def get_logs(self, level: Optional[LogLevel] = None, limit: int = 100) -> List[LogEntry]:
        """Get log entries, optionally filtered by level."""
        logs = self.logs
        if level:
            logs = [l for l in logs if l.level == level]
        return logs[-limit:]
    
    def get_errors(self) -> List[LogEntry]:
        """Get all error logs."""
        return [l for l in self.logs if l.level == LogLevel.ERROR]
    
    def get_warnings(self) -> List[LogEntry]:
        """Get all warning logs."""
        return [l for l in self.logs if l.level == LogLevel.WARNING]
    
    def clear_logs(self):
        """Clear all logs."""
        self.logs = []
    
    # -------------------------------------------------------------------------
    # SESSION INFO
    # -------------------------------------------------------------------------
    
    def get_summary(self) -> Dict[str, Any]:
        """Get session summary."""
        return {
            "session_id": self.state.session_id,
            "start_time": self.state.start_time.isoformat(),
            "files_processed": len(self.state.files_processed),
            "operations": len(self.state.operations_performed),
            "errors": len(self.get_errors()),
            "warnings": len(self.get_warnings()),
            "current_file": self.state.current_file,
        }
    
    def export_logs(self) -> str:
        """Export logs as JSON."""
        return json.dumps([l.to_dict() for l in self.logs], indent=2)
    
    def reset(self):
        """Reset session state."""
        self.state.reset()
        self.info("Session state reset")
