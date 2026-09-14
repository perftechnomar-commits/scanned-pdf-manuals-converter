"""Session-local, credential-free processing measurements."""
from contextvars import ContextVar
from datetime import datetime, timezone
import time

ACTIVE = ContextVar('processing_monitor', default=None)


class RunMonitor:
    def __init__(self, notify=lambda text: None, **metadata):
        self.notify = notify
        self.started = time.monotonic()
        self.stage = 'Preparation / reconciliation'
        self.rows = []
        self.retries = 0
        self.waits = {'Retry wait': 0.0, 'Request pacing': 0.0}
        self.metadata = metadata
        self.started_at = datetime.now(timezone.utc).isoformat()
        self.token = ACTIVE.set(self)

    def call(self, stage, function, *args, **kwargs):
        previous = self.stage
        self.stage = stage
        self.notify(stage)
        start = time.monotonic()
        status = 'Returned'
        try:
            return function(*args, **kwargs)
        except BaseException:
            status = 'Interrupted / failed'
            raise
        finally:
            self.rows.append({'Stage': stage, 'Seconds': round(time.monotonic()-start, 2), 'Status': status})
            self.stage = previous
            self.notify(previous)

    def finish(self, status):
        elapsed = time.monotonic()-self.started
        ACTIVE.reset(self.token)
        measured = sum(row['Seconds'] for row in self.rows)
        return dict(self.metadata, started_at=self.started_at, status=status,
                    elapsed_seconds=round(elapsed, 2), retries=self.retries,
                    waits=self.waits.copy(), stages=self.rows + [{
                        'Stage': 'Other processing / reconciliation',
                        'Seconds': round(max(0, elapsed-measured), 2), 'Status': status}])


def monitored_sleep(seconds, retry=False):
    monitor = ACTIVE.get()
    kind = 'Retry wait' if retry else 'Request pacing'
    if monitor is not None:
        monitor.retries += int(retry)
        monitor.notify(f'{monitor.stage} — {kind.lower()}: {seconds:.1f}s; retries: {monitor.retries}')
    started = time.monotonic()
    try:
        time.sleep(seconds)
    finally:
        if monitor is not None:
            monitor.waits[kind] += time.monotonic()-started
