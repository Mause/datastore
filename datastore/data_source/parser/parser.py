# from pprint import pprint
from dateutil.parser import parse as parse_date
from itertools import chain
import datetime
import json
import re

import xlrd

EXCLUDED_SHEET_NAMES = {
    'Contents'
}


def split(line):
    bits = line.split(', ')
    bits = chain.from_iterable(
        bit.split(':')
        for bit in bits
    )

    bits = map(str.strip, bits)

    return list(filter(bool, bits))


def parse_sheet_header(header):
    _, _, rest = header[0].split(' ', 2)

    data_name, *meta = split(rest.lower())

    data = {
        'name': data_name.title(),
        'meta': set(meta)
    }

    possibles_byes = [bit for bit in meta if bit.startswith('by')]
    if possibles_byes:
        data['by'] = possibles_byes[0]
        data['meta'].remove(possibles_byes[0])

    return data


def clean(lst):
    return list(filter(None, lst))


def to_html_table(rows):
    return '<table>{}</table>'.format('\n'.join(
        '<tr>{}</tr>'.format(''.join(
            '<td>{}</td>'.format(cell)
            for cell in row
        ))
        for row in rows
    ))


def remove_invalid_chars(row):
    return [
        str(cell).replace(b'\xe2\x80\x93'.decode(), '-')
        for cell in row
    ]


def parse_release_date(row):
    # slice the crap off the first cell
    bit = row[0][12:].replace('(Canberra time)', 'GMT+11')
    return parse_date(bit)


def parse_sheet(sheet):
    data = {}

    rows = list(sheet._cell_values)[1:-3]

    data.update({
        'id': rows.pop(0)[0].split()[0],
        'release_date': parse_release_date(rows.pop(0)),
        'meta': parse_sheet_header(rows.pop(0))
    })
    rows.pop(0)

    with open('sheets/{}.html'.format(sheet.name), 'w') as fh:
        fh.write(to_html_table(rows))

    # with open('sheets/{}.csv'.format(sheet.name), 'w') as fh:
    #     fh.write('\n'.join(
    #         '<>'.join(map(str, row))
    #         for row in rows
    #     ))

    sections = []

    # cheat for the moment
    row = remove_invalid_chars(rows[0])
    if re.match(r'\d{4}-\d{2}', row[1]):
        print(sheet.name, 'is by the months')
        headers = remove_invalid_chars(rows.pop(0)[1:])

    else:
        print(sheet.name, 'is layered')

        headers = []

        parents = remove_invalid_chars(rows.pop(0)[1:])
        children = remove_invalid_chars(rows.pop(0)[1:])
        for parent, sub_header in zip(parents, children):
            if parent:
                current_parent_header = parent

            if sub_header:
                headers.append('{} - {}'.format(
                    current_parent_header, sub_header
                ))
            else:
                headers.append(parent)

    data['scales'] = remove_invalid_chars(rows.pop(0)[1:])

    while rows:
        row = rows.pop(0)

        if len(clean(row)) == 1:
            # gotta new section
            if row[0].isupper():  # primary header
                section_name = row[0].title()
                if section_name.startswith('Rse'):
                    section_name = 'R.S.E.' + section_name[3:]

                sections.append((section_name, {}))

            else:
                # subheader, ignore
                pass

        else:
            # try:
                sections[-1]

                location, *values = row

                assert len(headers) == len(values)

                sections[-1][1][location] = dict(zip(
                    headers, values
                ))

            # except IndexError as e:
            #     print(e)
            #     return
                # print(sections)
                # raise

    # from pprint import pprint

    # sections is a list of tuples, so yeah
    data['sections'] = dict(sections)
    filename = 'sheets/{}.json'.format(sheet.name)
    with open(filename, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, indent=4, cls=SetSerializer)
    # pprint(sections)
    # pprint([
    #     [
    #         str(cell).encode('utf-8')
    #         for cell in row
    #     ]
    #     for row in rows
    # ])


class SetSerializer(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, set):
            return list(obj)

        elif isinstance(obj, datetime.datetime):
            return obj.timestamp()

        return super().default(obj)


def read_file(filename):
    doc = xlrd.open_workbook(filename)

    return [
        parse_sheet(sheet)
        for sheet in doc.sheets()
        # for sheet in [doc.sheets()[1]]
        if re.match(r'^Table \d+$', sheet.name)
        # if sheet.name not in EXCLUDED_SHEET_NAMES and
        # sheet.name.startswith('Table')
    ]


def main():
    read_file('time series and multiple victimisation.xls')

if __name__ == '__main__':
    main()
