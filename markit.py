from pptx import Presentation
import re


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

	def deletebySlideID(self,slide_id):

		slide = self.doc.slides.get(slide_id)
		index = self.doc.slides.index(slide)
		self.delete(index)

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
							# this detects the slide title
							seenOne = True
							slide_dict["title"] = paras[i].text
							slide_dict["id"] = str(slide.slide_id)
							element_dict = {}
							element_dict["key"] = "Title"
							element_dict["value"] = paras[i].text
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

						elif " Image" in paras[i].text:
							# this detected the slide type
							slide_dict["type"] = paras[i].text
							element_dict = {}
							element_dict["key"] = "Slide Type"
							element_dict["value"] = paras[i].text
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

						elif ":" in paras[i].text.lower() and len(paras[i].text) < 40:
							# this is detected as a key value entity
							element_dict = {}
							element_dict["key"] = paras[i].text.split(":")[0]
							element_dict["value"] = paras[i].text.split(":")[1]
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

						elif len(paras[i].text) > 40:
							# this is detected as a paragraph
							desc_count += 1
							element_dict = {}
							element_dict["key"] = "Description "+str(desc_count)
							element_dict["value"] = paras[i].text
							element_dict["pid"] = i
							slide_dict["elements"].append(element_dict)

			if "id" not in slide_dict:
				slide_dict["id"] = str(slide.slide_id)
				slide_dict["type"] = "Single Slide"
				slide_dict["elements"] = []
				slide_dict["title"] = "Intro Slide "+str(slideid)
				blueprint["Intro Slide "+str(slideid)] = [slide_dict]

			else:
				if slide_dict["title"] not in blueprint:
					blueprint[slide_dict["title"]] = []

				blueprint[slide_dict["title"]].append(slide_dict)


		return blueprint

	



	def stripPPT(self,selectedstuff):
		'''
		Function to remove all unnecessary sections from ppt
		and keep only the selected slides, and content,
		do so by sending a list of slide IDs and removing all
		slides not in that list.
		Similarly by sending all the content ids, and removing
		everything not in that list.
		'''

		slide_order = list(map(int,selectedstuff["slide_order"].split("_")))

		slide_ids_to_keep = []
		text_ids_to_keep = {}

		blueprint = self.blueprint()
		
		for key in selectedstuff.keys():

			if key[0] == "A":
				# keys that start with A_ represent sections to keep
				# and their values contain slide_IDs
				slide_ids_to_keep.append(int(selectedstuff[key]))

			elif key[0] == "B":
				# keys that start with B_ represent paragraphs to keep
				# and their values contain the actual values of the paragraph
				# format: B_slideid_textid_textkey
				value_split = key.split("_")
				slide_id = value_split[1]
				text_id = value_split[2]
				text_key = value_split[3]

				# add slideid key to dict
				if slide_id not in text_ids_to_keep:
					text_ids_to_keep[slide_id] = {}

				# create a heirarchieal datastructure to represent the text
				# to keep and what value to store
				text_ids_to_keep[slide_id][text_id] = [text_key,selectedstuff[key]]

		# remove unselected text from slides 
		for slide_id in text_ids_to_keep:
			for shape in self.doc.slides.get(int(slide_id)).shapes:
				if hasattr(shape, "text"):
					paras = shape.text_frame.paragraphs
					
					i = 0
					# i tracks the paragraph index in the textbox
					# we keep track of which paragraph index are selected
					for para in paras:
						if ((str(i) not in text_ids_to_keep[slide_id]) and 
							(para.text != "")):
							
							# actual delete paragraph ops
							p = para._p
							p.getparent().remove(p)

						elif (str(i) in text_ids_to_keep[slide_id]):
							
							value = text_ids_to_keep[slide_id][str(i)][1]

							if ":" in para.text:
								runs = para.runs
								# we assume there is only single run
								if len(runs) > 1:
									
									runs[0].text = para.text
									runs[0].text = runs[0].text.replace("_x000B_","")
									
									for j in range(1,len(runs)):
										runs[j].text = ""

								index = para.runs[0].text.index(":")
								paras[i].runs[0].text = paras[i].runs[0].text[:index+1] + value
								
							else:
								runs = para.runs
								if len(runs) > 1:
									runs[0].text = para.text
									for j in range(1,len(runs)):
										runs[j].text = ""
								
								paras[i].runs[0].text = value
						i+=1



		# remove slides that are not selected
		for section in blueprint.keys():
			for slide in blueprint[section]:
				if int(slide["id"]) not in slide_ids_to_keep:
					self.deletebySlideID(int(slide["id"]))

		i = 0
		for slide_id in slide_order:

			slide = self.doc.slides.get(slide_id)
			index = self.doc.slides.index(slide)

			self.move_slide(index, i)
			i+=1

	def move_slide(self, old_index, new_index):
		xml_slides = self.doc.slides._sldIdLst  # pylint: disable=W0212
		slides = list(xml_slides)
		xml_slides.remove(slides[old_index])
		xml_slides.insert(new_index, slides[old_index])

					
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



