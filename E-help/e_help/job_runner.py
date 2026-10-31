"""Tk-independent background runner; queue events are the UI boundary."""
import queue
import threading

from .analysis import CancelledError, check_cancel


class JobRunner:
    """One worker per window; event queue is the sole worker-to-UI channel."""
    def __init__(self):
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.thread = None

    @property
    def busy(self):
        return self.thread is not None and self.thread.is_alive()

    def start(self, operation, *, accept_completed=False):
        if self.busy:
            return False
        while not self.events.empty():
            self.events.get_nowait()
        self.cancel_event = threading.Event()
        # Non-daemon: let cancellation close/save resources even if the parent exits.
        self.thread = threading.Thread(target=self._run, args=(operation, accept_completed), daemon=False)
        self.thread.start()
        return True

    def _run(self, operation, accept_completed=False):
        try:
            result = operation(self.cancel_event, lambda text: self.events.put(("progress", text)))
            if not accept_completed:
                check_cancel(self.cancel_event)
            self.events.put(("done", result))
        except CancelledError as exc:
            self.events.put(("cancelled", str(exc)))
        except Exception as exc:
            self.events.put(("error", f"{type(exc).__name__}: {exc}"))

    def cancel(self):
        self.cancel_event.set()
