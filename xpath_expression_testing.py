

def escape_string_for_xpath(s):
    if '"' in s and "'" in s:
        return 'concat(%s)' % ", '\"',".join('"%s"' % x for x in s.split('"'))
    elif '"' in s:
        return "'%s'" % s
    return '"%s"' % s


myobj = "Testing a contraction inside string. It's OK if this doesn't work."

s = []
s.append(myobj)
escaped_title = escape_string_for_xpath(s[0])
print(escaped_title)
