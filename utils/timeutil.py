from time import strftime, localtime, time


def get_cur_time_str(date_format="%Y-%m-%d %H:%M:%S"):
    return strftime(date_format, localtime(time()))
