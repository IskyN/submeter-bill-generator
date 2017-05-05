from getpass import getpass
from io import StringIO

print("here")
s = StringIO()
p = getpass("don't enter a password. ", stream=s)
print(p, s.getvalue())
