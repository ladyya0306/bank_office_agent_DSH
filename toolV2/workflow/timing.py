"""Measured intervals in the original task JSON; gaps are explicitly unclassified."""
from contextlib import contextmanager
from datetime import datetime, timezone
import time


def now():
    return time.time()


def stamp(value):
    return datetime.fromtimestamp(value, timezone.utc).isoformat(timespec='milliseconds')


def add(task, name, seconds, started, ended):
    clock = task.setdefault('timing', {'stages': {}, 'wait_ids': []})
    stage = clock['stages'].setdefault(name, {'seconds': 0, 'count': 0, 'first_start': stamp(started)})
    stage['seconds'] = round(stage['seconds'] + max(0, seconds), 6)
    stage['count'] += 1
    stage['last_end'] = stamp(ended)


def _task_started(clock, current):
    """Earliest known task time, including pre-upgrade task payloads."""
    value = clock.get('first_started_at')
    if isinstance(value, (int, float)):
        return value
    starts = []
    for stage in clock.get('stages', {}).values():
        try:
            starts.append(datetime.fromisoformat(stage['first_start']).timestamp())
        except (KeyError, TypeError, ValueError):
            pass
    value = min(starts) if starts else current
    clock['first_started_at'] = value
    return value


@contextmanager
def measure(task, name):
    wall, tick = now(), time.perf_counter()
    try:
        yield
    finally:
        add(task, name, time.perf_counter() - tick, wall, now())


def begin(task, wait=None):
    current = now()
    clock = task.setdefault('timing', {'stages': {}, 'wait_ids': []})
    task_started = _task_started(clock, current)
    previous = clock.get('last_end')
    gap = max(0, current - previous) if previous else 0
    waited = 0
    overlap = 0
    valid_wait = False
    if isinstance(wait, dict) and wait.get('id') and wait['id'] not in clock['wait_ids']:
        # Telemetry comes from the native question controller, never from model tool parameters.
        milliseconds = wait.get('milliseconds')
        wait_start, wait_end = wait.get('started_at'), wait.get('ended_at')
        hosted_interval = (isinstance(wait_start, (int, float)) and isinstance(wait_end, (int, float))
                           and task_started <= wait_start <= wait_end <= current)
        if hosted_interval:
            waited = wait_end - wait_start
            valid_wait = True
        elif previous is not None and isinstance(milliseconds, (int, float)) and 0 <= milliseconds <= gap * 1000:
            waited = milliseconds / 1000
            wait_start, wait_end = current - waited, current
            valid_wait = True
        if valid_wait:
            add(task, 'user_confirmation_wait', waited, wait_start, wait_end)
            clock['wait_ids'].append(wait['id'])
            clock['cancelled_waits'] = clock.get('cancelled_waits', 0) + int(wait.get('outcome') == 'cancelled')
            if previous is not None:
                overlap = max(0, min(wait_end, current) - max(wait_start, previous))
    if gap > overlap:
        phase = 'mapping_interval_unclassified' if clock.get('last_status') == 'needs_mapping' else 'other_interval_unclassified'
        add(task, phase, gap - overlap, previous, current)
    return current


def end(task, started):
    current = now()
    add(task, 'tool_calls', current - started, started, current)
    task['timing'].update(last_end=current, last_status=task.get('status'))


def summary(task):
    return {'unit': 'seconds', 'stages': task.get('timing', {}).get('stages', {}),
            'cancelled_waits': task.get('timing', {}).get('cancelled_waits', 0),
            'note': 'tool_calls 包含程序子阶段，不能重复相加。用户等待计原生提问调用的实测等待，含问题显示和返回开销；它可能与并发程序阶段重叠，因此各阶段 seconds 也不能直接相加。未分类间隔只扣除了与等待区间重叠的部分。位置整理间隔含模型、会话及传输，未细分。seconds 是累计秒数，first_start/last_end 为记录边界，不表示中间连续运行。未接入计时的历史区间、意外终止前未记录的程序耗时无法还原。'}
