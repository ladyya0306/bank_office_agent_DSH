"""Read a selected business workspace in SQLite read-only mode."""
import argparse
import json
from pathlib import Path
import sqlite3

SECTIONS = {'facts': 'fact', 'answers': 'fill_decision', 'outputs': 'fill_op',
            'events': 'event', 'tasks': 'office_v2_task', 'rules': 'template_rule'}


def main():
    parser = argparse.ArgumentParser(description='只读查看办公数据库；结果可能包含客户信息，请勿公开上传。')
    parser.add_argument('--work', type=Path, required=True, help='业务工作区目录')
    parser.add_argument('--section', choices=SECTIONS, default='tasks')
    parser.add_argument('--limit', type=int, default=10)
    args = parser.parse_args()
    path = (args.work / 'db/workflow.db').resolve(strict=True)
    with sqlite3.connect(path.as_uri() + '?mode=ro', uri=True) as db:
        db.row_factory = sqlite3.Row
        db.execute('PRAGMA query_only=ON')
        table = SECTIONS[args.section]
        if not db.execute('SELECT 1 FROM sqlite_master WHERE type=? AND name=?', ('table', table)).fetchone():
            print('该库还没有这一类记录。')
            return
        rows = db.execute(f'SELECT * FROM "{table}" ORDER BY rowid DESC LIMIT ?',
                          (max(1, min(args.limit, 100)),)).fetchall()
        print(json.dumps([dict(row) for row in rows], ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
