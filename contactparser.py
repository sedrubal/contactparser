#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
# PYTHON_ARGCOMPLETE_OK

"""
a command line tool to convert microsofts .contact files into a csv or json
"""

from __future__ import print_function

__author__ = "sedrubal"
__license__ = "GPLv3"
__url__ = "https://github.com/sedrubal/contactparser"

import sys
from bs4 import BeautifulSoup
import collections
import json
import csv
import argparse
try:
    from argcomplete import autocomplete
except ImportError:
    pass


def parse_args():
    """
    Parse arguments and autocomplete
    :return: the parsed args
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('files',
                        action='store',
                        nargs='+',
                        metavar='file',
                        type=argparse.FileType('r'),
                        help=".contact files, - for stdin")
    parser.add_argument('-v', '--verbose',
                        action='count',
                        dest='verbosity',
                        default=False,
                        help="prints debug messages to stderr, -vv for more)")
    parser.add_argument('-o', '--output',
                        action='store',
                        dest='output_file',
                        metavar='file',
                        type=argparse.FileType('w'),
                        default='-',
                        help="the output file, - for stdout")
    parser_json_grp = parser.add_argument_group('output: json')
    parser_json_grp.add_argument('--json',
                                 action='store_const',
                                 const='json',
                                 dest='output_format',
                                 help="output format is json")
    parser_json_grp.add_argument('--pretty',
                                 action='store_true',
                                 dest='json_pretty',
                                 required=False,
                                 help='Make json pretty and not compact')
    parser_csv_grp = parser.add_argument_group('output: csv')
    parser_csv_grp.add_argument('--csv',
                                action='store_const',
                                const='csv',
                                dest='output_format',
                                help="output format is csv")
    parser_csv_grp.add_argument('--csv-dialect',
                                action='store',
                                dest='csv_dialect',
                                choices=['excel', 'excel-tab', 'unix'],
                                default='unix',
                                help="the csv dialect")

    if 'autocomplete' in globals():
        autocomplete(parser)

    args = parser.parse_args()

    # parser.add_mutually_exclusive_group did not work so check this manually
    if args.output_format == 'csv' and args.json_pretty:
        parser.exit("[!] '--pretty' is only for json output")
        sys.exit(1337)
    if args.output_format == 'json' and '--csv-dialect' in sys.argv:
        parser.exit("[!] '--csv-dialect' is only for csv output")
        sys.exit(1337)

    if not args.output_format and args.output_file:
        if args.output_file.name.endswith('.csv'):
            verbose_print("setting output format to csv (from fileextension)",
                          args.verbosity, 2)
            args.output_format = 'csv'
        elif args.output_file.name.endswith('.csv'):
            verbose_print("setting output format to json (from fileextension)",
                          args.verbosity, 2)
            args.output_format = 'json'
    if not args.output_format and args.json_pretty:
        verbose_print("setting output format to json (from --pretty)",
                      args.verbosity, 2)
        args.output_format = 'json'
    if not args.output_format and '--csv-dialect' in sys.argv:
        verbose_print("setting output format to csv (from --csv-dialect)",
                      args.verbosity, 2)
        args.output_format = 'csv'

    if not args.output_format:
        verbose_print("setting output format to json (default)",
                      args.verbosity, 2)
        args.output_format = 'json'

    return args


def verbose_print(message, verbosity, on_verbosity_level=1):
    """
    prints if the current verbosity level is >= the on_verbosity_level
    """
    if verbosity >= on_verbosity_level:
        print('[i] ', message, file=sys.stderr)


def print_debug_id(element, args):
    """
    prints the ElementID of an element when being very verbose
    """
    element_id = element.get('c:ElementID')
    if element_id is not None:
        verbose_print("│   ├── processing id '%s'" % element_id,
                      args.verbosity, 2)


def safe_find_all(element, value):
    """
    calls findAll(value) of element if element is not None.
    Otherwise an empty list ([]) will be returned
    """
    if element is not None:
        return element.findAll(value)
    return []


def safe_get_text(element, default=''):
    """
    calls getText() of element if element is not None.
    Otherwise default will be returned
    """
    if element is not None:
        return element.getText()
    return default


def parse_contact_file(contact_file, args):
    """
    Parses the given .contact file and returns a dict containing the
    contact information
    """
    verbose_print("opening '%s'" % contact_file.name, args.verbosity)
    contact_xml = BeautifulSoup(contact_file.read(), 'xml')
    con = {}

    # Names
    # it's possible that one contact has many names in namecollection.
    verbose_print("├── parsing names", args.verbosity, 2)
    con['Name'] = []
    for name in safe_find_all(contact_xml.NameCollection, 'Name'):
        print_debug_id(name, args)
        if not name.get('xsi:nil'):
            nam = {}
            nam['FormattedName'] = safe_get_text(name.FormattedName)
            nam['FamilyName'] = safe_get_text(name.FamilyName)
            nam['GivenName'] = safe_get_text(name.GivenName)
            con['Name'].append(nam)

    # Email addresses
    # it's possible that one contact has many email adresses in
    # EmailAddressCollection
    verbose_print("├── parsing email adresses", args.verbosity, 2)
    con['Email'] = []
    for email in safe_find_all(contact_xml.EmailAddressCollection,
                               'EmailAddress'):
        print_debug_id(email, args)
        if not email.get('xsi:nil'):
            emaddr = {}
            emaddr['Address'] = safe_get_text(email.Address)
            emaddr['Labels'] = []
            for label in safe_find_all(email.LabelCollection, 'Label'):
                emaddr['Labels'].append(safe_get_text(label))
                con['Email'].append(emaddr)

    return con


def make_contact_csv_compatible(contact, verbosity):
    """
    makes a parsed contact csv compatible as
    it must be storageable in a 2D table:
        - each field has to get it's own name
        - multiple names have to be merged in one name
        - email labels will get lost
    :return: an ordered dict containing the simplified values of contact
    """
    verbose_print("├── converting contact", verbosity, 2)
    # simplified contacts have this fields:
    new_contact = collections.OrderedDict()
    new_contact['FormattedName'] = ''
    new_contact['GivenName'] = ''
    new_contact['FamilyName'] = ''
    new_contact['Email-Preferred'] = ''
    new_contact['Email-1'] = ''
    new_contact['Email-2'] = ''
    new_contact['Email-3'] = ''
    new_contact['Email-4'] = ''

    def copy_if_valid(dict1, dict2, fieldname):
        """
        copies the value of fieldname in dict1 to dict2 if it is a valid entry
        """
        if dict1[fieldname].strip():
            if dict2[fieldname]:
                verbose_print(
                    "│   ├── overwriting '{field}' with '{new}'."
                    "'{old}' will get lost".format(field=fieldname,
                                                   new=dict1[fieldname],
                                                   old=dict2[fieldname]),
                    verbosity, 1)
            dict2[fieldname] = dict1[fieldname].strip()

    for nam in contact['Name']:
        copy_if_valid(nam, new_contact, 'FamilyName')
        copy_if_valid(nam, new_contact, 'GivenName')
        copy_if_valid(nam, new_contact, 'FormattedName')

    email_count = 0
    for emaddr in contact['Email']:
        if 'preferred' in [label.lower() for label in emaddr['Labels']]:
            new_contact['Email-Preferred'] = emaddr['Address'].strip()
        else:
            email_count += 1
            new_contact['Email-%d' % email_count] = emaddr['Address'].strip()
            if email_count > 4:
                verbose_print(
                    "│   ├── Too many email addresses. '%s' will get lost" %
                    emaddr['Address'].strip(), verbosity, 1)
    return new_contact


def write_json(contacts, output_file, pretty=False, verbosity=0):
    """
    makes a json out of the contacts dict and
    writes it to output_file (an open file descriptor)
    """
    verbose_print("generating json", verbosity, 1)
    if pretty:
        json_string = json.dumps(contacts,
                                 sort_keys=True,
                                 indent=4,
                                 separators=(',', ': '))
    else:
        json_string = json.dumps(contacts)
    print(json_string, file=output_file)


def write_csv(contacts, output_file, dialect='unix', verbosity=0):
    """
    makes a csv out of the contacts dict and
    writes it to output_file (an open file descriptor)
    """
    verbose_print("generating csv", verbosity, 1)
    new_contacts = []
    for contact in contacts:
        # make contacts csv compatible
        new_contacts.append(make_contact_csv_compatible(contact, verbosity))

    writer = csv.DictWriter(output_file,
                            fieldnames=new_contacts[0].keys(),
                            dialect=dialect)

    writer.writeheader()
    for contact in new_contacts:
        writer.writerow(contact)


def save_output(contacts, args):
    """
    saves the contacts to file
    """

    # write to file or to stdout if no output file is given
    if args.output_format == 'json':
        write_json(contacts,
                   args.output_file,
                   args.json_pretty,
                   args.verbosity)
    elif args.output_format == 'csv':
        write_csv(contacts,
                  args.output_file,
                  args.csv_dialect,
                  args.verbosity)
    else:
        print("[!] can't generate a %s - not implemented" %
              args.output_format,
              sys.stderr)
        sys.exit(1337)


def main():
    """
    The main function parses each given xml file and
    saves the output into a file
    """
    args = parse_args()

    contacts = []

    for contact_file in args.files:
        contacts.append(parse_contact_file(contact_file, args))

    save_output(contacts, args)


if __name__ == '__main__':
    main()
    exit(0)
