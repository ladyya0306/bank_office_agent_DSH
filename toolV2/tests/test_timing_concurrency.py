from unittest.mock import patch

from workflow import timing


def test_host_wait_survives_concurrent_last_end_and_deduplicates_id():
    task = {}
    with patch.object(timing, 'now', return_value=100):
        started = timing.begin(task)
    with patch.object(timing, 'now', return_value=101):
        timing.end(task, started)
    # Another agent's stage advanced the shared task clock during the native
    # question.  The host interval itself remains valid task telemetry.
    task['timing']['last_end'] = 105
    event = {'id': 'native-parallel-1', 'milliseconds': 4000,
             'started_at': 102, 'ended_at': 106, 'outcome': 'answered'}
    with patch.object(timing, 'now', return_value=106):
        timing.begin(task, event)
    with patch.object(timing, 'now', return_value=108):
        timing.begin(task, event)
    stages = timing.summary(task)['stages']
    assert stages['user_confirmation_wait']['seconds'] == 4
    assert stages['user_confirmation_wait']['count'] == 1
    # Only [105,106] overlaps the first gap.  The duplicate event contributes
    # no waiting time; its later three-second gap remains honestly unclassified.
    assert stages['other_interval_unclassified']['seconds'] == 3


def test_answered_within_one_clock_tick_is_still_one_question_batch():
    task = {}
    with patch.object(timing, 'now', return_value=100):
        timing.begin(task)
    event = {'id': 'instant-answer', 'started_at': 101, 'ended_at': 101,
             'milliseconds': 0, 'outcome': 'answered'}
    with patch.object(timing, 'now', return_value=102):
        timing.begin(task, event)
        timing.begin(task, event)
    waiting = timing.summary(task)['stages']['user_confirmation_wait']
    assert waiting['seconds'] == 0
    assert waiting['count'] == 1
