from pptx import Presentation


prs = Presentation('Mr. Naval_The Quarry Stones.pptx')

def delete_slides(presentation, index):
	xml_slides = presentation.slides._sldIdLst  
	slides = list(xml_slides)
	print(len(xml_slides))
	xml_slides.remove(slides[index])

for slide in prs.slides:
	# print(slide.slide_id)
	for shape in slide.shapes:
		if hasattr(shape, "text"):
			# shape.text_frame.paragraphs[0].runs[0].text
			paras = shape.text_frame.paragraphs
			for i in range(len(paras)):
				r = paras[i].runs
				
				# print(i,paras[i].text)
				for j in range(len(r)):

					# if i == 5 and j == 0:
					# 	shape.text_frame.paragraphs[i].runs[j].text += "Added"
					print(i,j,r[j].text)
	print()

delete_slides(prs,0)

# prs.save("edited.pptx")