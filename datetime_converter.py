from datetime import time, datetime


def to_datetime(time: time):
    return datetime(1, 1, 1, time.hour, time.minute, time.second)
