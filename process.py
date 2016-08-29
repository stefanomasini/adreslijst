import xlrd
import vobject
import sys


def readRowsWithXlrd(filename, sheetIndex):
    wbook = xlrd.open_workbook(filename)
    sheet = wbook.sheet_by_index(sheetIndex)
    numRows = sheet.nrows
    numCols = sheet.ncols

    for rowIdx in xrange(numRows):
        row = sheet.row(rowIdx)
        yield [row[cellIdx].value for cellIdx in xrange(numCols)]


def iterAddresses(filepath):
    for idx, row in enumerate(readRowsWithXlrd(filepath, 0)):
        if idx == 0:
            headers = [h.strip().lower() for h in row]
        else:
            yield dict(zip(headers, [(v.strip() if v else '') for v in row]))


def generateVCards(addresses):
    for address in addresses:
        card = vobject.vCard()

        # Name
        card.add('n')
        card.n.value = vobject.vcard.Name( family=address.get('achternaam', ''), given=address.get('naam', '') )
        card.add('fn')
        card.fn.value = (address.get('naam', '') + ' ' + address.get('achternaam', '')).strip()

        # Email
        card.add('email')
        card.email.value = address.get('email', '')
        card.email.type_param = 'INTERNET'

        # Org (vader/moeder van ...)
        if address.get('kind', ''):
            ge = address.get('ge', '')
            card.add('org')
            prefix = 'ouder van'
            if ge and ge.lower() in 'mv':
                prefix = 'vader van' if ge.lower() == 'v' else 'moeder van'
            card.org.value = ['%s %s' % (prefix, address.get('kind'))]

        # Phone
        if address.get('telefoonnummer', ''):
            card.add('tel')
            card.tel.value = address.get('telefoonnummer', '')
            card.tel.type_param = 'CELL'

        # Address
        if address.get('adres', '') or address.get('postcode', '') or address.get('stad', ''):
            card.add('adr')
            data = {}
            for key, value in (
                    ('street', address.get('adres', '')),
                    ('code', address.get('postcode', '')),
                    ('city', address.get('stad', '')),
                ):
                if value:
                    data[key] = value
            card.adr.value = vobject.vcard.Address(**data)
            card.adr.type_param = 'HOME'

        yield card


def iterEmailAddresses(addresses):
    for address in addresses:
        parts = [address.get('naam', '')]
        if address.get('kind', ''):
            prefix = 'ouder van'
            ge = address.get('ge', '')
            if ge and ge.lower() in 'mv':
                prefix = 'vader van' if ge.lower() == 'v' else 'moeder van'
            parts.extend([prefix, address.get('kind')])
        yield '%s <%s>,' % (' '.join(parts), address.get('email'))


if __name__ == '__main__':
    filepath = sys.argv[1]
    addresses = list(iterAddresses(filepath))
    vcfFilepath = filepath.rsplit('.', 1)[0] + '.vcf'
    print >> open(vcfFilepath, 'w'), ''.join(card.serialize().strip() for card in generateVCards(addresses))
    mlFilepath = filepath.rsplit('.', 1)[0] + ' mailing list.txt'
    print >> open(mlFilepath, 'w'), '\n'.join(iterEmailAddresses(addresses))
