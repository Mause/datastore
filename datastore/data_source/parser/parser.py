import xlrd
from itertools import chain

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


# def determine_type()


def parse_sheet(sheet):
    rows = list(sheet._cell_values)[3:-3]

    data = {}
    data['meta'] = parse_sheet_header(rows.pop(0))
    data['headers'] = [cell for cell in rows.pop(0) if cell]

    with open('sheets/{}.csv'.format(sheet.name), 'w') as fh:
        fh.write('\n'.join(
            '<>'.join(map(str, row))
            for row in rows
        ))

    with open('sheets/{}.html'.format(sheet.name), 'w') as fh:
        fh.write(to_html_table(rows))

    sections = []

    print(data['meta']['name'], data['headers'])
    print(remove_invalid_chars(rows[0]))

    # cheat for the moment
    rows = rows[2:]

    while rows:
        row = rows.pop(0)

        if len(clean(row)) == 1:
            # gotta new sections
            section_name = row[0].title()
            if section_name.startswith('Rse'):
                section_name = 'R.S.E.' + section_name[3:]

            sections.append((section_name, {}))

        else:
            try:
                sections[-1]

                location, *values = row

                sections[-1][1][location] = dict(zip(
                    data['headers'], values
                ))

            except IndexError:
                return
                # print(sections)
                # raise

    from pprint import pprint

    # sections is a list of tuples, so yeah
    sections = dict(sections)
    pprint(sections)
    # pprint([
    #     [
    #         str(cell).encode('utf-8')
    #         for cell in row
    #     ]
    #     for row in rows
    # ])


def read_file(filename):
    doc = xlrd.open_workbook(filename)

    return [
        parse_sheet(sheet)
        for sheet in doc.sheets()
        # for sheet in [doc.sheets()[1]]
        if sheet.name not in EXCLUDED_SHEET_NAMES and
        sheet.name.startswith('Table')
    ]


def main():
    read_file('time series and multiple victimisation.xls')

if __name__ == '__main__':
    main()
