from markit import Document

d = Document('Mr. Naval_The Quarry Stones.pptx')

d.deletebyTitle("Ombianca")
d.save("deleted.pptx")