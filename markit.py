from pptx import Presentation


class Document(object):
	"""docstring for Document"""
	def __init__(self, docname):
		super(Document, self).__init__()
		
		self.docname = docname
		self.doc = Presentation(self.docname)

	def delete(self,slideno):
		xml_slides = self.doc.slides._sldIdLst  
		slides = list(xml_slides)
		xml_slides.remove(slides[slideno])

	def deletebyTitle(self,title):
		slidestoDelete = []
		i = 0
		for slide in self.doc.slides:
			for shape in slide.shapes:
				if hasattr(shape, "text"):
					paras = shape.text_frame.paragraphs
					if paras[0].text == title:
						slidestoDelete.append(i)
						break
			i+=1

		print("Slides to Delete:",slidestoDelete)

		for j in range(len(slidestoDelete)):
			# re-adjusting index to accomodate reindexing of slide list
			newindex = slidestoDelete[j] - j
			self.delete(newindex)


	def editKeyValue(self,slideid,key,value):

	def editParagraph(self,slideid,paraid,content):
		for shape in prs.slides[slideid].shapes:
			if hasattr(shape, "text"):
				paras = shape.text_frame.paragraphs
				if len(paras[paraid].text) > 40 :
					# qualifies as paragraph


	def save(self,name):
		self.doc.save(name)




# TODO: 
# Adding tags to ppt elements, for template creation
# PPT Versioning for various clients, and history of what was sent when and to whom
# 



