import urllib.parse
import ast
import codecs

try:
    with codecs.open('app.js', 'r', encoding='utf-8') as f:
        content = f.read()
except UnicodeDecodeError:
    with codecs.open('app.js', 'r', encoding='latin-1') as f:
        content = f.read()

# Since it's Javascript, python's ast won't work. We need to use esprima or something.
# We don't have node. We can just use python to extract basic info or search for unclosed brackets.
