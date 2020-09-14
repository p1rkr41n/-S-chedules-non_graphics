def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

def excel_column_number(name):
    """Excel-style column name to number, e.g., A = 1, Z = 26, AA = 27, AAA = 703."""
    n = 0
    for c in name:
        n = n * 26 + 1 + ord(c) - ord('A')
    return n

def test (name, number):
    for n in [0, 1, 2, 3, 24, 25, 26, 27, 702, 703, 704, 2708874, 1110829947]:
        a = name(n)
        n2 = number(a)
        a2 = name(n2)
        print ("%10d  %-9s  %s" % (n, a, "ok" if a == a2 and n == n2 else "error %d %s" % (n2, a2)))