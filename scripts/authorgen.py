#!/usr/bin/env python
import xlrd

def load(filename, sheet=0):
    '''
    Load spread sheet file, return array of dictionaries
    '''
    wb = xlrd.open_workbook(filename)
    s = wb.sheets()[0]
    keys = [c.value.lower() for c in s.row(0)] # column names

    ret = list()
    for irow in range(1, s.nrows):
        d = {k:c.value for k,c in zip(keys, s.row(irow))}
        #print irow,d['fname'],d['lname']
        ret.append(d)
    return ret

def latex(registrants):
    '''
    \author{K.S.~Babu\footnotemark[1]\footnotetext[1]{Convener}}
    \affiliation{Oklahoma State University}
    '''
    ret = list()
    for person in registrants:
        inits = person['fname'][0].upper() + '.'
        if person['mi']:
            inits += person['mi'].upper() + '.'
        d = dict(person, inits=inits, footnote='')
        s = r'''\author{%(inits)s~%(lname)s%(footnote)s}
\affiliation{%(affiliation)s}''' % d
        ret.append(s)
    return '\n'.join(ret)

if '__main__' == __name__:
    import sys
    d = load(sys.argv[1])
    print latex(d)
