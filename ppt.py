from pptx import Presentation


prs = Presentation('Jumbo 3.pptx')

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
				runs = paras[i].runs
				
				if len(runs) > 1:
					runs[0].text = paras[i].text
				# print(i,paras[i].text)
					for j in range(1,len(runs)):
						runs[j].text = ""
						# if i == 5 and j == 0:
						# 	shape.text_frame.paragraphs[i].runs[j].text += "Added"
						# print(i,j,runs[j].text)
				for j in range(len(runs)):
					print(i,j,runs[j].text)

	print()

# delete_slides(prs,0)

# prs.save("edited.pptx")