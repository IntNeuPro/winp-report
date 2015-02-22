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
        #print irow,d
        ret.append(d)
    return ret

def sort_affiliations(affils):
    'Return sorted tuples of (nick,long name) of affiliations'
    # maybe do something smarter here.
    return sorted(affils.items())

def latex(registrants):
    '''
    \author{K.S.~Babu\footnotemark[1]\footnotetext[1]{Convener}}
    \affiliation{Oklahoma State University}
    '''

    affils = dict()
    authors = list()

    def make_affil(affil):
        ascii_affil = affil.decode('unicode_escape').encode('ascii','ignore')
        nick = ascii_affil.translate(None, ' .-,_/')
        return affils.setdefault(ascii_affil, nick)

    for person in registrants:
        inits = person['fname'][0].upper() + '.'
        if person['mi']:
            inits += person['mi'].upper() + '.'

        extra_fn = ""
        if person['convenor']:  # watch out for the alternative spelling!
            extra_fn += r'\footnotemark[1]'
        if person['organizing']: 
            extra_fn += r'\footnotemark[2]'

        affil_nick = make_affil(person['affiliation'])
        d = dict(person, inits=inits, footnote=extra_fn, affil=affil_nick)
        s = r'''\author{%(inits)s~%(lname)s%(footnote)s}
\affiliation{\%(affil)s}''' % d
        authors.append(s)

    ret = list()
    ret.append(r"\footnotetext[1]{Convener}")
    ret.append(r"\footnotetext[2]{Organizer}")

    for inst,nick in sort_affiliations(affils):
        ret.append(r'''\newcommand{\%s}{%s}''' % (nick, inst)) # must match usage above
        ret.append(r'''\affiliation{\%s}''' % (nick,))

    ret.extend(authors)
    return '\n'.join(ret)

if '__main__' == __name__:
    import sys
    d = load(sys.argv[1])
    try:
        outfile = sys.argv[2]
        outfp = open(outfile, 'w')
    except IndexError:
        outfp = sys.stdout

    outfp.write(latex(d))
    outfp.close()

