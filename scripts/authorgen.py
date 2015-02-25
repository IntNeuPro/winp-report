#!/usr/bin/env python
import re
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

def affiliator(affils):
    def make_affil(affil):
        ascii_affil = affil.decode('unicode_escape').encode('ascii','ignore')
        nick = ascii_affil.translate(None, ' .-,_/')
        return affils.setdefault(ascii_affil, nick)
    return make_affil

def initials(person):
    inits = person['fname'][0].upper() + '.'
    if person['mi']:
        inits += person['mi'].upper() + '.'
    return inits

def affiliations(person):
    'Return affiliations for a person as a list'
    return [x.strip() for x in re.split('[/]', person['affiliation'])]


def authblk(registrants):
    'Return an authblk author list'
    ret = [r'''
\renewcommand*{\thefootnote}{\fnsymbol{footnote}}
\footnotetext[1]{Convenor}
\newcommand\ConvenorMark{\footnotemark[1]}
\footnotetext[2]{Organizer}
\newcommand\OrganizerMark{\footnotemark[2]}
\renewcommand*{\thefootnote}{\arabic{footnote}}
''']
    convenor_fn = r'\protect\ConvenorMark'
    organizer_fn = r'\protect\OrganizerMark'


    affils = set()
    for person in registrants:
        affils.update(affiliations(person))
    affils = list(affils)
    affils.sort()

    affil_offset = 1

    for person in registrants:
        extra_fn = ""
        extra_ret = list()

        if person['convenor']:  # watch out for the alternative spelling!
            extra_fn += convenor_fn

        if person['organizing']: 
            extra_fn += organizer_fn

        ind = [str(affils.index(a)+affil_offset) for a in affiliations(person)]
        ind = ','.join(sorted(ind))
        d = dict(person, inits=initials(person), 
                 footnote = extra_fn,
                 ind = ind)
        ret.append(r'\author[%(ind)s]{%(inits)s~%(lname)s%(footnote)s}' % d)
        ret.extend(extra_ret)

    for count, affil in enumerate(affils):
        ret.append(r'\affil[%d]{\mbox{%s}}' % (count+affil_offset, affil))

    return '\n'.join(ret)

if '__main__' == __name__:
    import sys
    d = load(sys.argv[1])
    try:
        outfile = sys.argv[2]
        outfp = open(outfile, 'w')
    except IndexError:
        outfp = sys.stdout

    #latex = revtex(d)
    latex = authblk(d)

    outfp.write(latex)
    outfp.close()

