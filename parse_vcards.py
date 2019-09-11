#!/usr/bin/env python
"""Parse vcard stream from stdin, output a CSV for consumption by Office365

"""
import sys
import csv
from collections import namedtuple, OrderedDict


vcard_collection = []

FIELDS_CSV = 'ExternalEmailAddress Name FirstName LastName StreetAddress City StateorProvince PostalCode Phone MobilePhone Pager HomePhone Company Title OtherTelephone Department CountryOrRegion Fax Initials Notes Office Manager'.split()


def new_vcard(suffix):
    global the_vcard
    the_vcard = {a: '' for a in FIELDS_CSV}
    the_vcard['Manager'] = 'parse_vcard.py'

def vcard_version(suffix):
    global the_vcard
    the_vcard['version'] = suffix

def vcard_name(suffix):
    global the_vcard
    try:
        ln, fn, _ = suffix.split(';', 2)
    except Exception:
        ln = suffix
        fn = ''
    the_vcard['LastName'] = ln
    the_vcard['FirstName'] = fn

def vcard_fullname(suffix):
    global the_vcard
    the_vcard['Name'] = suffix

def vcard_phone(suffix):
    global the_vcard
    typ, val = suffix.split(':', 1)
    _, typ = typ.split('=')
    if typ == 'home':
        the_vcard['HomePhone'] = val
    elif typ == 'CELL':
        the_vcard['MobilePhone'] = val
    elif typ == 'work':
        the_vcard['Phone'] = val
    else:
        the_vcard['OtherTelephone'] = val

def vcard_company(suffix):
    global the_vcard
    the_vcard['Company'] = suffix

def vcard_title(suffix):
    global the_vcard
    the_vcard['Title'] = suffix

def vcard_email(suffix):
    global the_vcard
    typ, val = suffix.split(':', 1)
    the_vcard['ExternalEmailAddress'] = val

def vcard_addr(suffix):
    global the_vcard
    typ, val = suffix.split(':', 1)
    _, typ = typ.split('=', 1)
    try:
        _, _, full, street, city, state, zipcode, country = val.split(';')
    except ValueError:
        full = street = city = state = zipcode = country = ''
    the_vcard['PostalCode'] = zipcode
    the_vcard['StateorProvince'] = state
    the_vcard['City'] = city
    the_vcard['StreetAddress'] = street
    the_vcard['CountryOrRegion'] = country

def vcard_note(suffix):
    global the_vcard
    the_vcard['Notes'] = suffix

def vcard_department(suffix):
    global the_vcard
    the_vcard['Department'] = suffix

def vcard_end(suffix):
    global the_vcard
    global vcard_collection
    the_vcard.pop('version')
    c = Contact(**the_vcard)
    vcard_collection.append(c)

def parse_vcard(line):
    fun = None
    pre = None
    for p, f in prefixes.items():
        if not line.startswith(p):
            continue
        fun = f
        pre = p.strip()
    if not fun:
        print("cannot parse>", line)
        return -1
    suffix = line[len(pre):].strip()
    fun(suffix)



prefixes = {
'BEGIN:VCARD': new_vcard,
'VERSION:': vcard_version,
'N:': vcard_name,
'FN:': vcard_fullname,
'TEL;': vcard_phone,
'ORG:': vcard_company,
'TITLE:': vcard_title,
'EMAIL;': vcard_email,
'ADR;': vcard_addr,
'NOTE:': vcard_note,
'X-DEPARTMENT:': vcard_department,
'END:VCARD': vcard_end,
        }

Contact = namedtuple('Contact', FIELDS_CSV)



for line in sys.stdin.readlines():
    parse_vcard(line)

with open(sys.argv[1], 'w') as f:
    c = csv.writer(f)
    c.writerow(FIELDS_CSV)
    c.writerows([[getattr(r, v) for v in FIELDS_CSV] for r in vcard_collection])


