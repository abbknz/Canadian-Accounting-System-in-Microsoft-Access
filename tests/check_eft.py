"""Checks a CPA-005 file made by the program against the record layout of the bank's guide."""
import sys

lines = open(sys.argv[1], encoding='cp1252', newline='').read().split('\r\n')
assert lines[-1] == '', 'the file must end with a line break'
lines = lines[:-1]
routing = lines.pop(0) if lines[0].startswith('$$') else None
bad = [i for i, l in enumerate(lines) if len(l) != 1464]
a, z, cs = lines[0], lines[-1], lines[1:-1]
pay, total = [], 0
for c in cs:
    assert c[0] == 'C'
    for k in range(6):
        seg = c[24 + 240 * k: 264 + 240 * k]
        if seg.strip():
            pay.append(seg)
            total += int(seg[3:13])
print('routing record:', routing)
print('records:', len(lines), '| wrong length:', bad, '| record counts in sequence:', [int(l[1:10]) for l in lines] == list(range(1, len(lines) + 1)))
print('header: type', a[0], '| client', a[10:20], '| file number', a[20:24], '| created', a[24:30], '| centre', a[30:35], '| currency', a[55:58])
print('payments:', len(pay), '| same client and file number on every record:', all(l[10:24] == a[10:24] for l in lines))
s = pay[0]
print('first payment: code', s[0:3], '| cents', s[3:13], '| date', s[13:19], '| institution+transit', s[19:28], '| account', repr(s[28:40]), '| zeros', s[40:65] == '0' * 25,
      '| short name', repr(s[65:80]), '| payee', repr(s[80:110].rstrip()), '| originator', repr(s[110:140].rstrip()), '| client', s[140:150], '| reference', repr(s[150:169].rstrip()),
      '| zeros', s[169:178] == '0' * 9, '| sundry', repr(s[190:205].rstrip()))
print('trailer: type', z[0], '| total cents', int(z[46:60]), '(sum of payments', total, ') | count', int(z[60:68]), '| zero filled', set(z[68:]) == {'0'} and z[24:46] == '0' * 22)
print('all numeric fields are digits:', all(s[3:28].isdigit() for s in pay))
