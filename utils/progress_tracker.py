"""
Progress Tracker Module

Provides progress tracking functionality for long-running operations.
"""

import time
from typing import Optional, Callable, Any
from dataclasses import dataclass
from threading import Lock


@dataclass
class ProgressState:
    """Current progress state"""
    current: int
    total: int
    percentage: float
    message: str
    elapsed_seconds: float
    estimated_remaining_seconds: Optional[float]
    items_per_second: float


class ProgressTracker:
    """
    Tracks progress of multi-step operations.
    Thread-safe for use with background processing.
    """
    
    def __init__(
        self,
        total: int,
        callback: Optional[Callable[[ProgressState], None]] = None,
        update_interval: float = 0.1  # Minimum seconds between callbacks
    ):
        """
        Initialize progress tracker.
        
        Args:
            total: Total number of items to process
            callback: Optional callback function for progress updates
            update_interval: Minimum interval between callback invocations
        """
        self.total = max(1, total)  # Avoid division by zero
        self.callback = callback
        self.update_interval = update_interval
        
        self._current = 0
        self._message = ""
        self._start_time = time.time()
        self._last_callback_time = 0
        self._lock = Lock()
    
    def update(self, current: Optional[int] = None, message: str = ""):
        """
        Update progress.
        
        Args:
            current: Current item number (or None to increment by 1)
            message: Progress message
        """
        with self._lock:
            if current is not None:
                self._current = current
            else:
                self._current += 1
            
            if message:
                self._message = message
            
            self._maybe_notify()
    
    def increment(self, amount: int = 1, message: str = ""):
        """
        Increment progress by a given amount.
        
        Args:
            amount: Amount to increment by
            message: Progress message
        """
        with self._lock:
            self._current += amount
            if message:
                self._message = message
            self._maybe_notify()
    
    def set_message(self, message: str):
        """Set progress message without changing current value."""
        with self._lock:
            self._message = message
            self._maybe_notify()
    
    def set_total(self, total: int):
        """Update the total count."""
        with self._lock:
            self.total = max(1, total)
    
    def _maybe_notify(self):
        """Notify callback if enough time has passed."""
        if not self.callback:
            return
        
        current_time = time.time()
        
        # Always notify on completion or if interval has passed
        if self._current >= self.total or \
           current_time - self._last_callback_time >= self.update_interval:
            
            self._last_callback_time = current_time
            state = self._get_state()
            
            try:
                self.callback(state)
            except Exception:
                pass  # Don't let callback errors break progress tracking
    
    def _get_state(self) -> ProgressState:
        """Get current progress state."""
        elapsed = time.time() - self._start_time
        percentage = (self._current / self.total) * 100
        
        # Calculate rate and ETA
        if elapsed > 0 and self._current > 0:
            items_per_second = self._current / elapsed
            remaining_items = self.total - self._current
            
            if items_per_second > 0:
                estimated_remaining = remaining_items / items_per_second
            else:
                estimated_remaining = None
        else:
            items_per_second = 0
            estimated_remaining = None
        
        return ProgressState(
            current=self._current,
            total=self.total,
            percentage=percentage,
            message=self._message,
            elapsed_seconds=elapsed,
            estimated_remaining_seconds=estimated_remaining,
            items_per_second=items_per_second
        )
    
    def get_state(self) -> ProgressState:
        """Get current progress state (thread-safe)."""
        with self._lock:
            return self._get_state()
    
    @property
    def current(self) -> int:
        """Get current progress value."""
        with self._lock:
            return self._current
    
    @property
    def percentage(self) -> float:
        """Get current percentage complete."""
        with self._lock:
            return (self._current / self.total) * 100
    
    @property
    def is_complete(self) -> bool:
        """Check if progress is complete."""
        with self._lock:
            return self._current >= self.total
    
    def reset(self, total: Optional[int] = None):
        """
        Reset the progress tracker.
        
        Args:
            total: New total (or None to keep current)
        """
        with self._lock:
            if total is not None:
                self.total = max(1, total)
            self._current = 0
            self._message = ""
            self._start_time = time.time()
            self._last_callback_time = 0


class MultiStageProgressTracker:
    """
    Tracks progress across multiple stages.
    Each stage has a weight that determines its contribution to overall progress.
    """
    
    def __init__(
        self,
        stages: dict,  # {stage_name: weight}
        callback: Optional[Callable[[str, float, str], None]] = None
    ):
        """
        Initialize multi-stage progress tracker.
        
        Args:
            stages: Dictionary mapping stage names to weights
            callback: Callback(stage_name, overall_percentage, message)
        """
        self.stages = stages
        self.callback = callback
        
        # Normalize weights
        total_weight = sum(stages.values())
        self.normalized_weights = {
            name: weight / total_weight 
            for name, weight in stages.items()
        }
        
        # Calculate cumulative weights for each stage start
        self.stage_starts = {}
        cumulative = 0
        for name in stages:
            self.stage_starts[name] = cumulative
            cumulative += self.normalized_weights[name]
        
        self._current_stage = None
        self._stage_progress = 0
    
    def start_stage(self, stage_name: str):
        """
        Start a new stage.
        
        Args:
            stage_name: Name of the stage to start
        """
        if stage_name not in self.stages:
            raise ValueError(f"Unknown stage: {stage_name}")
        
        self._current_stage = stage_name
        self._stage_progress = 0
        self._notify("")
    
    def update_stage(self, percentage: float, message: str = ""):
        """
        Update progress within current stage.
        
        Args:
            percentage: Stage completion percentage (0-100)
            message: Progress message
        """
        self._stage_progress = min(100, max(0, percentage))
        self._notify(message)
    
    def _notify(self, message: str):
        """Notify callback with overall progress."""
        if not self.callback or not self._current_stage:
            return
        
        # Calculate overall percentage
        stage_start = self.stage_starts[self._current_stage]
        stage_weight = self.normalized_weights[self._current_stage]
        stage_contribution = stage_weight * (self._stage_progress / 100)
        
        overall_percentage = (stage_start + stage_contribution) * 100
        
        self.callback(self._current_stage, overall_percentage, message)
    
    def complete_stage(self):
        """Mark current stage as complete."""
        self._stage_progress = 100
        self._notify("Complete")
    
    @property
    def overall_percentage(self) -> float:
        """Get overall percentage complete."""
        if not self._current_stage:
            return 0
        
        stage_start = self.stage_starts[self._current_stage]
        stage_weight = self.normalized_weights[self._current_stage]
        stage_contribution = stage_weight * (self._stage_progress / 100)
        
        return (stage_start + stage_contribution) * 100
