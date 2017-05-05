"""
Various attempts from
http://stackoverflow.com/questions/3160699/python-progress-bar
Best: Pascal @ http://stackoverflow.com/a/21840062
"""

# from time import sleep
# from sys import stdout
#
# toolbar_width = 40
#
# # setup toolbar
# stdout.write('[' + ' ' * toolbar_width + ']')
# stdout.flush()
# stdout.write("\b" * (toolbar_width+1))  # return to start of line, after '['
#
# for i in range(toolbar_width):
#     sleep(0.2)  # do real work here
#     # update the bar
#     stdout.write("-")
#     stdout.flush()
#
# stdout.write("\n")

################################################################################

# import time, sys
#
# def update_progress(progress):
#     """Displays or updates a console progress bar
#     Accepts a float between 0 and 1. Any int will be converted to a float.
#     A value under 0 represents a 'halt'.
#     A value at 1 or bigger represents 100%
#
#     :type progress: float
#     :return: None
#     """
#     bar_length = 10  # Modify this to change the length of the progress bar
#     status = ""
#     if isinstance(progress, int):
#         progress = float(progress)
#     if not isinstance(progress, float):
#         progress = 0
#         status = "error: progress var must be float\r\n"
#     if progress < 0:
#         progress = 0
#         status = "Halt...\r\n"
#     if progress >= 1:
#         progress = 1
#         status = "Done...\r\n"
#     block = int(round(bar_length*progress))
#     text = "\rPercent: [{0}] {1}% {2}".format(
#         "#"*block + "-"*(bar_length-block), progress*100, status)
#     sys.stdout.write(text)
#     sys.stdout.flush()
#
#
# # update_progress test script
# print("progress : 'hello'")
# update_progress("hello")
# time.sleep(1)
#
# print("progress : 3")
# update_progress(3)
# time.sleep(1)
#
# print("progress : [23]")
# update_progress([23])
# time.sleep(1)
#
# print()
# print("progress : -10")
# update_progress(-10)
# time.sleep(2)
#
# print()
# print("progress : 10")
# update_progress(10)
# time.sleep(2)
#
# print()
# print("progress : 0->1")
# for i in range(101):
#     time.sleep(0.01)
#     update_progress(i/100.0)
#
# print()
# print("Test completed")
# time.sleep(10)

################################################################################

from time import sleep
from sys import stdout


def update_progress_bar(percent: float):
    if percent == 1:
        print("Progress: {:3.1%}".format(percent))
    else:
        print("Progress: {:3.1%}\r".format(percent), end="")
        # stdout.flush()

for i in range(5):
    sleep(0.1)
    update_progress_bar(i/5)
update_progress_bar(1)
