import json
import sys

import pymupdf as fitz

path = sys.argv[1]
doc = fitz.open(path)
words = []
text_parts = []
for pi, page in enumerate(doc):
    text_parts.append(page.get_text())
    for w in page.get_text("words"):
        words.append({"str": w[4], "x": w[0], "y": w[1], "page": pi + 1})
print(json.dumps({"words": words, "text": "\n".join(text_parts)}))
