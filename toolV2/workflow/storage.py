"""Task progress and cached source rows in the existing office database."""
import hashlib
import json
import os
from contextlib import contextmanager
from pathlib import Path


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def initialise(conn):
    conn.executescript('''
      CREATE TABLE IF NOT EXISTS office_v2_task(id TEXT PRIMARY KEY, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS office_v2_cache(id TEXT PRIMARY KEY, payload TEXT NOT NULL);
    ''')


def get(conn, table, key):
    assert table in ('office_v2_task', 'office_v2_cache')
    row = conn.execute(f'SELECT payload FROM {table} WHERE id=?', (key,)).fetchone()
    return json.loads(row[0]) if row else None


def put(conn, table, key, value):
    assert table in ('office_v2_task', 'office_v2_cache')
    conn.execute(f'INSERT INTO {table}(id,payload) VALUES(?,?) ON CONFLICT(id) DO UPDATE SET payload=excluded.payload',
                 (key, json.dumps(value, ensure_ascii=False)))
    conn.commit()


def inside(work, name):
    p = (work / name).resolve(strict=True)
    if not p.is_relative_to(work) or not p.is_file():
        raise ValueError(f'文件不在本工作区内：{name}')
    return p


@contextmanager
def work_lock(work):
    # OS releases this lock if the process exits. No persistent permission token.
    p = work / 'db' / '.toolv2-working'
    p.parent.mkdir(exist_ok=True)
    with p.open('a+b') as f:
        if p.stat().st_size == 0:
            f.write(b'0'); f.flush()
        f.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise RuntimeError('本工作区已有填报正在执行，请等待该次结果；未重复启动。') from exc
        try:
            yield
        finally:
            f.seek(0)
            if os.name == 'nt':
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(f, fcntl.LOCK_UN)
