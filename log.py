import datetime


def logStatus(logFilePath, logMsg, need_time_stamp=True):
    """
    Write the logMsg to the log file, with option to include timestamp or not.
    :param logFilePath: the path you would like to save the log file
    :param logMsg: Message in string
    :param need_time_stamp: True or False
    :return:
    """
    file = open(logFilePath, "a")
    if logMsg == "\n":
        file.write("\n")
    else:
        if need_time_stamp:
            time_format = "%Y-%m-%d %H:%M:%S"
            file.write(f"{logMsg} @ {datetime.datetime.now().strftime(time_format)}")
        else:
            file.write(logMsg)
        file.write("\n")
    file.close()