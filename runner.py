from markit import Document
import pprint

d = Document('Mr. Naval_The Quarry Stones.pptx')

bp = d.blueprint()

pprint.pprint(bp)

# d.deletebyTitle("Ombianca")
# d.save("ready.pptx")