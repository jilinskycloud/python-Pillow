############################
#     Written BY UMER      #
#  umershah81@hotmail.com  #
############################

from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import PIL
import xlrd
import os
import requests
import shutil
import glob
import math

class EditImage:

	def __init__(self):
		self.title_font = ImageFont.truetype('SimHei.ttf', 8, encoding="utf-8")
		self.rows = ""
		self.columns = ""
		self.r1 = ""
		self.r2 = ""
		self.ch_list = [None] * 6
		self.p_img = '' #Image.open('2.jpg')
		#self.I1 = ImageDraw.Draw(self.p_img)
		self.loc = "" #("20210329.xls")
		self.wb = '' #xlrd.open_workbook(self.loc)
		self.sheet = ''#self.wb.sheet_by_index(0)
		self.indi = 0
		self.rd1 = 0
		self.rd2 = 0
		self.rd = 0

	def downImage(self, image_url, filename):
		## Set up the image URL and filename
		print("Down URL :: ", image_url)
		print("This is the file name ::", filename)
		filenamePath = filename +".jpg"
		r = requests.get(image_url, stream = True)
		# Check if the image was retrieved successfully
		if r.status_code == 200:
			# Set decode_content value to True, otherwise the downloaded image file's size will be zero.
			r.raw.decode_content = True
			# Open a local file with wb ( write binary ) permission.
			with open("downloads/" + filenamePath,'wb') as f:
				shutil.copyfileobj(r.raw, f)
			print('Image sucessfully Downloaded: ',filenamePath)
			return filenamePath
		else:
			print('Image Couldn\'t be retreived')

	def readText(self):
		self.r1 = str(self.sheet.cell_value(0, 1)) + ":" + self.ch_list[1] + "  " + str(self.sheet.cell_value(0, 2)) + ":" + self.ch_list[2]
		self.r2 = str(self.sheet.cell_value(0, 3)) + ":" + self.ch_list[3] + "     " + str(self.sheet.cell_value(0, 4)) + ":" + str(int(self.ch_list[4]))

	def resize(self, filenamePath):
		newsize = (132,170)
		self.p_img = "downloads/"+filenamePath
		self.p_img = Image.open(self.p_img)
		self.p_img = self.p_img.resize(newsize)
		# Save the edited image
		self.p_img.save("downloads/" + filenamePath, quality=100, subsampling=0)

	def createbg(self):
		self.readText()
		bg = Image.new('RGB', (132, 27), color = 'white')
		image_editable = ImageDraw.Draw(bg)
		#image_editable.text((2,2), self.r1, (255, 255, 255), font=self.title_font, fill='black')
		#image_editable.text((2,12), self.r2, (255, 255, 255), font=self.title_font, fill='black')
		image_editable.text((2,2), self.r1, font=self.title_font, fill='black')
		image_editable.text((2,12), self.r2, font=self.title_font, fill='black')
		bg.save('white_bnr.png', quality=100, subsampling=0)

	def concat(self, filenamePath):
		list_im = ["downloads/" + filenamePath, 'white_bnr.png']
		imgs = [Image.open(im) for im in list_im]
		width_of_new_image = 132
		height_of_new_image = 197
		new_im = Image.new('RGB', (width_of_new_image, height_of_new_image))
		new_pos = 0
		print("My file name is :: ", filenamePath)
		for im in imgs:
			new_im.paste(im, (0, new_pos))
			new_pos += im.size[1] #position for the next image
		new_im.save('edited/' + filenamePath, quality=100, subsampling=0) #change the filename if you want

	def readFileAndPutInArray(self, r):
		#for r in range(2):
		for c in range(self.columns):
			self.ch_list[c] = self.sheet.cell_value(r, c)

	def processImages(self):
		#in file rows = 70 and columns = 6
		self.loc = input("Please Enter the .XLS File name ::")
		self.rows = int(input("Please enter Number of Rows :: "))
		self.columns = int(input("Please enter Number of Colmns :: "))
		self.wb = xlrd.open_workbook(self.loc)
		self.sheet = self.wb.sheet_by_index(0)
		for r in range(1,self.rows):
			self.readFileAndPutInArray(r)
			print("This is the Array :: ", self.ch_list)
			filenamePath = self.downImage(self.ch_list[5], str(int(self.ch_list[0])))
			self.resize(filenamePath)
			self.createbg()
			self.concat(filenamePath)

	def comb(self):
		try:
			files = glob.glob("./edited/*.jpg")
			#result = Image.new("RGB", (793, 1122))
			lnI = math.ceil(len(files)/25)
			#print("Round IT :: ", lnI)
			#print("Round IT :: ", round(lnI))
			for xvp in range(round(lnI)):
				result = Image.new("RGB", (668, 1025), color = 'white') # (793, 1122))
				for i in range(self.indi,self.indi+25): #len(files)): #file in enumerate(files):
					path = os.path.expanduser(files[i])
					img = Image.open(path)
					img.thumbnail((132, 197), Image.ANTIALIAS)
					x = self.rd // 5 * 134
					y = self.rd % 5 * 207
					w, h = img.size
					#print('pos {0},{1} size {2},{3}'.format(x, y, w, h))
					result.paste(img, (x, y, x + w, y + h))
					result.save(os.path.expanduser('./combine_img/' + str(xvp) + '_image.png'), quality=100, subsampling=0)
					self.indi = i
					self.rd=self.rd+1
					x = x+20
					y = y+50
				self.rd=0

		except:
			pass		

def main(obj):
	obj.processImages()
	obj.comb()
	print("DONE ...!")

if __name__ == '__main__':
	obj = EditImage()
	main(obj)
