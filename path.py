import os
from urllib.parse import urlparse
import urllib.parse

p= "&#46;∖../../../..//.//.//.//..//"
p2 = os.path.normpath(p)
print (p)
print (p2)

query = 'https://www.google.com?search=query'
o = urllib.parse.quote(query)
print(o)