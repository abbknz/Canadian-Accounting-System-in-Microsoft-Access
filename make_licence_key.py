"""Makes the licence key for one customer and records it in customers.csv.

    python make_licence_key.py "Customer Name"                 key without an end date
    python make_licence_key.py "Customer Name" 2027-12-31      key that stops working after that day

The customer types the same name and the key into About / Licence. Capitals, spaces and punctuation in the
name do not matter. The key is worked out from licence_secret.txt, the secret that build_logic.py puts into
the program: keep that file private and back it up - with another secret, the keys already given out stop working.
"""
import csv, datetime, hashlib, os, re, sys

here = os.path.dirname(os.path.abspath(__file__))


def make_key(name, expiry='00000000'):
    clean = ''.join(ch for ch in name.upper() if ch in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')
    if not clean:
        sys.exit('The name needs at least one letter or digit (A-Z, 0-9).')
    secret = open(os.path.join(here, 'licence_secret.txt')).read().strip()
    h = hashlib.sha256((secret + '|' + clean + '|' + expiry).encode('ascii')).hexdigest().upper()
    return 'KA-%s-%s-%s-%s-%s' % (expiry, h[0:5], h[5:10], h[10:15], h[15:20])


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    name = sys.argv[1].strip()
    expiry = '00000000'
    if len(sys.argv) > 2:
        expiry = datetime.datetime.strptime(sys.argv[2], '%Y-%m-%d').strftime('%Y%m%d')
    key = make_key(name, expiry)
    version = re.search(r'AppVersion = "([^"]+)"', open(os.path.join(here, 'modAccounting.bas'), encoding='cp1252').read()).group(1)
    log = os.path.join(here, 'customers.csv')
    new = not os.path.exists(log)
    with open(log, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if new:
            w.writerow(['issued', 'customer', 'ends', 'key', 'program version'])
        w.writerow([datetime.date.today().isoformat(), name, sys.argv[2] if len(sys.argv) > 2 else '', key, version])
    print('Licensed to :', name)
    print('Licence key :', key)
    print('Ends        :', sys.argv[2] if len(sys.argv) > 2 else 'no end date')
