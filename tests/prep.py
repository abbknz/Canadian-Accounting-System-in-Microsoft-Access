"""Shared by the tests: makes a freshly opened copy ready for use (licence accepted, activated, setup marked done)."""
import hashlib, os

SECRET = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'licence_secret.txt')


def make_key(name, expiry='00000000'):
    clean = ''.join(ch for ch in name.upper() if ch in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')
    h = hashlib.sha256((open(SECRET).read().strip() + '|' + clean + '|' + expiry).encode('ascii')).hexdigest().upper()
    return 'KA-%s-%s-%s-%s-%s' % (expiry, h[0:5], h[5:10], h[10:15], h[15:20])


def ready(app, setup=True):
    app.Run('AcceptLicence')
    r = app.Run('TryActivate', 'Test Company', make_key('Test Company'))
    assert (r[0] if isinstance(r, tuple) else r) == 'OK', r
    if setup:
        app.Run('SkipSetup')
    app.Run('StartUp')
