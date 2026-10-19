"""Makes the DASHBOARD_PASS_HASH value for a hosted copy of the tool (run:  py tools/make_password_hash.py).

Only the hash goes into the host's settings; the password itself is never stored anywhere."""
import getpass
import sys

from werkzeug.security import generate_password_hash

pw = getpass.getpass("New dashboard password: ")
if len(pw) < 10:
    sys.exit("Use at least 10 characters.")
if pw != getpass.getpass("Type it again: "):
    sys.exit("The two entries differ.")
print("\nSet this as a variable on the host (copy the whole line after the = sign):\n")
print("DASHBOARD_PASS_HASH=" + generate_password_hash(pw))
