"""Presentation helpers shared by the desktop UI and its tests."""
def duration(seconds):
    seconds = max(0, int(seconds))
    return f'{seconds // 60}m {seconds % 60:02d}s'


def estimate(completed, total, elapsed):
    return elapsed / completed * (total - completed) if completed >= 3 and total > completed else None
