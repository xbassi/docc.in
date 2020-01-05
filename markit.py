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


	def blueprint(self):

		blueprint = {}
		slideid = 0
		for slide in self.doc.slides:
			slideid+=1
			slide_dict = {}
			seenOne = False
			desc_count = 0
			for shape in slide.shapes:
				if hasattr(shape, "text"):
					paras = shape.text_frame.paragraphs
					slide_dict["type"] = "Details"
					slide_dict["elements"] = []
					for i in range(len(paras)):
						if seenOne == False and len(paras[i].text) > 1 and len(paras[i].text) < 40: 
							seenOne = True
							slide_dict["title"] = paras[i].text
							slide_dict["id"] = paras[i].text

						elif " Image" in paras[i].text:
							slide_dict["type"] = paras[i].text

						elif ":" in paras[i].text.lower() and len(paras[i].text) < 40:
							element_dict = {}
							element_dict["key"] = paras[i].text.split(":")[0]
							element_dict["value"] = paras[i].text.split(":")[1]
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

						elif len(paras[i].text) > 40:
							desc_count += 1
							element_dict = {}
							element_dict["key"] = "Description "+str(desc_count)
							element_dict["value"] = paras[i].text
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

			if "id" not in slide_dict:
				slide_dict["id"] = slideid
				slide_dict["title"] = "Intro Slide "+str(slideid)
				blueprint["Intro Slide "+str(slideid)] = slide_dict

			else:
				if slide_dict["id"] not in blueprint:
					blueprint[slide_dict["id"]] = []

				blueprint[slide_dict["id"]].append(slide_dict)



		return blueprint

	def editKeyValue(self,slideid,pid,key,value):

		for shape in self.doc.slides[slideid].shapes:
			if hasattr(shape, "text"):
				paras = shape.text_frame.paragraphs
				for i in range(len(paras)):
					r = paras[i].runs
					if ((":" in paras[i].text) and 
						(key.lower() in paras[i].text.lower())):

						index = paras[i].text.index(":")
						# we assume there is only single run
						paras[i].runs[0].text = paras[i].runs[0].text[:index+1] + value
						return True

		return False


	def editParagraph(self,slideid,paraid,content):
		for shape in self.doc.slides[slideid].shapes:
			if hasattr(shape, "text"):
				paras = shape.text_frame.paragraphs
				if len(paras[paraid].text) > 40 :
					# qualifies as paragraph
					pass

	def save(self,name):
		self.doc.save(name)




# TODO: 
# Adding tags to ppt elements, for template creation
# PPT Versioning for various clients, and history of what was sent when and to whom
# 



