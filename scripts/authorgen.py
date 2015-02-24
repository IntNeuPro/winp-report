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

def revtex(registrants):
    '''
    Return a revtex4 compatible author list
    '''

    authors = list()

    affils = dict()
    make_affil = affiliator(affils)

    for person in registrants:
        inits = initials(person)
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
        affil = person['affiliation']
        affils.add(affil)
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

        d = dict(person, inits=initials(person), 
                 footnote = extra_fn,
                 ind = affils.index(person['affiliation'])+affil_offset)

        a = r'\author[%(ind)d]{%(inits)s~%(lname)s%(footnote)s}' % d
        ret.append(a)
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

