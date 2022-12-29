#pip install pysimplegui
#pip install pillow
#pip install opencv-python
#pip install numpy
from PIL import Image
import os
import fitz  # fitz就是pip install PyMuPDF
from fpdf import FPDF
import numpy as np
import cv2
import PySimpleGUI as sg
from PyPDF2 import PdfFileReader, PdfFileWriter
import time
import shutil

layout = [
	[
        sg.FolderBrowse(
            button_text="拆分文件夹中文件",
            initial_folder=r"C:/Users/oristand/Desktop/zhouhui/change",
            target="folder_path"
        ),
        sg.InputText(key="folder_path"),
		sg.Button('splitPdf')
	],
	[sg.Combo((
		'fedex Wayfair识别运单号错误类似于（77572622）---不处理旋转---处理方案为cut_type:10识别',
		'默认ups,fedex,可裁剪---不需要处理旋转---处理方案为cut_type:36识别',
		'ups---需要处理旋转逆时针270---处理方案为cut_type:36识别',
		'wayfair入口 fedex（条码未到底部）---不需要处理旋转---处理方案为cut_type:10识别',
		'other入口 ups---不需要处理旋转---最下面有shipRush符号:处理方案为cut_type:36,底部加空白',
		'other入口 ups（四周空白，label在中间，慎用）---不需要处理旋转---处理方案为cut_type:36识别',
		'other入口 fedex（需要高度符合该label慎用）---需要处理旋转逆时针90---处理方案为cut_type:34识别',
		'wayfair 入口 PO:CS （慎用）---不需要处理旋转---处理方案为cut_type:9识别',
		'wayfair 入口 PO:CS （需要高度符合该label慎用）---不需要处理旋转---处理方案为cut_type:9识别',
		'amazon 入口 fedex （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:6识别',
		'单个pdf中只含一张的情况 fedex （需要高度符合该label慎用）---处理旋转90---处理方案为cut_type:34识别',
		'一个pdf好多张ups（第一页） （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:36识别',
		'一个pdf好多张ups（非第一页） （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:36识别',
		'fedex 下面留白太多---不处理旋转---处理方案为cut_type:34识别',
		'正常xdp无法识别',
		'other ups 四周空白中间有label',
		'other ups label在左侧，多页的第一页',
		'other ups label在左侧，多页的非第一页',
		'other ups label 逆时针90度，四周有空白',
	),
	'fedex Wayfair识别运单号错误类似于（77572622）---不处理旋转---处理方案为cut_type:10识别',key='cut_type',enable_events=True,auto_size_text=True,size=(100,100))],
	[sg.Image(filename='./img/1-1.png',key='image'),sg.Image(filename='./img/1-1.png',key='image1'),sg.Image(filename='./img/1-1.png',key='image2')],
	[
        sg.FolderBrowse(
            button_text="请选择需要处理pdf的文件夹",
            initial_folder=r"C:/Users/oristand/Desktop/zhouhui/labels",
            target="labels"
        ),
        sg.InputText("C:/Users/oristand/Desktop/zhouhui/labels",key="labels"),
		sg.Button('pdfDeal')
	],
	[sg.Text("label处理进度条:",visible=False,key='progress_bar_text'),sg.ProgressBar(1000, orientation='h', size=(80, 20), key='progress_bar',visible=False),sg.Text("0%:",visible=False,key='progress_bar_num')],
	[sg.Text("处理后的pdf缩略图,请仔细查看:",visible=False,key='img_deal_text')],
	[sg.Image(filename='',key='img_deal_0'),sg.Image(filename='',key='img_deal_1'),sg.Image(filename='',key='img_deal_2')],
]


def recursion_dir_all_file(path):
	'''
	:param path: 文件夹目录
	'''
	file_list = []
	for dir_path, dirs, files in os.walk(path):
		for file in files:
			file_path = os.path.join(dir_path, file)
			if "\\" in file_path:
				file_path = file_path.replace('\\', '/')
			file_list.append(file_path)
	for dir in dirs:
			file_list.extend(recursion_dir_all_file(os.path.join(dir_path, dir)))
	return file_list


def get_now_datetime():
    """
    @Description: 返回当前时间，格式为：年月日时分秒
    """
    return time.strftime('%Y-%m-%d',time.localtime(time.time() - 3600*16))
# PDF文件分割
def split_pdf(file_name,tag_dir,need_create_dir):
	try:
		read_file = file_name
		if os.path.splitext(file_name)[-1] != '.pdf':
			return
		file_name_without_suffix = os.path.basename(file_name)[0:-4]
		fp_read_file = open(read_file, 'rb')
		pdf_input = PdfFileReader(fp_read_file)  # 将要分割的PDF内容格式话
		page_count = pdf_input.getNumPages()  # 获取PDF页数
# 		print("this pdf has {} page".format(page_count))  # 打印页数
		last_parent_dir = tag_dir + need_create_dir + '/'
		if not os.path.exists(os.path.dirname(last_parent_dir)):
			os.makedirs(os.path.dirname(last_parent_dir))
		if page_count == 1:
		    shutil.copyfile(file_name, tag_dir + need_create_dir + '/' + file_name_without_suffix + '.pdf')
		    return
		for i in range(page_count):
			start_page = i
			end_page = i+1
			pdf_file = tag_dir + need_create_dir + '/' + file_name_without_suffix + '_part_' + str(start_page) + '.pdf'
			try:
# 				print(f' start {start_page} page -{end_page} page,save as {pdf_file}......')
				pdf_output = PdfFileWriter()  # 实例一个 PDF文件编写器
				for i in range(start_page, end_page):
					pdf_output.addPage(pdf_input.getPage(i))
				with open(pdf_file, 'wb') as sub_fp:
					pdf_output.write(sub_fp)
# 				print(f' finish {start_page} page -{end_page} page，save as {pdf_file}!')
			except IndexError:
				print(f' out of page')
	except Exception as e:
		print(e)

window = sg.Window('label处理小工具', layout, size=(900, 1000))

def cut_type_config():
	return {
		'fedex Wayfair识别运单号错误类似于（77572622）---不处理旋转---处理方案为cut_type:10识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':1},
		'默认ups,fedex,可裁剪---不需要处理旋转---处理方案为cut_type:36识别':{'degree':0,'format':[102,152],'save_width':800,'start_spot_x':0,'start_spot_y':0, 'end_spot_x':0,'end_spot_y':0,'image_sort':2},
		'ups---需要处理旋转逆时针270---处理方案为cut_type:36识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':3},
		'wayfair入口 fedex（条码未到底部）---不需要处理旋转---处理方案为cut_type:10识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':4},
		'other入口 ups---不需要处理旋转---最下面有shipRush符号:处理方案为cut_type:36,底部加空白':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':5},
		'other入口 ups（四周空白，label在中间，慎用）---不需要处理旋转---处理方案为cut_type:36识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':6},
		'other入口 fedex（需要高度符合该label慎用）---需要处理旋转逆时针90---处理方案为cut_type:34识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':7},
		'wayfair 入口 PO:CS （慎用）---不需要处理旋转---处理方案为cut_type:9识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':8},
		'wayfair 入口 PO:CS （需要高度符合该label慎用）---不需要处理旋转---处理方案为cut_type:9识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':9},
		'amazon 入口 fedex （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:6识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':10},
		'单个pdf中只含一张的情况 fedex （需要高度符合该label慎用）---处理旋转90---处理方案为cut_type:34识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':11},
		'一个pdf好多张ups（第一页） （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:36识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':12},
		'一个pdf好多张ups（非第一页） （需要高度符合该label慎用）---处理旋转270---处理方案为cut_type:36识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':13},
		'fedex 下面留白太多---不处理旋转---处理方案为cut_type:34识别':{'degree':3,'format':[102,154],'save_width':800,'start_spot_x':1351,'start_spot_y':391, 'end_spot_x':3930,'end_spot_y':4080,'image_sort':14},
		'正常xdp无法识别':{'degree':0,'format':[102,127],'save_width':800,'pdf_img_width':95,'start_spot_x':20,'start_spot_y':0, 'end_spot_x':2592,'end_spot_y':3240,'image_sort':15},
		'other ups 四周空白中间有label':{'degree':0,'format':[102,152],'save_width':800,'start_spot_x':171,'start_spot_y':403, 'end_spot_x':2448,'end_spot_y':3449,'image_sort':16},
		'other ups label在左侧，多页的第一页':{'degree':0,'format':[102,152],'save_width':800,'start_spot_x':1225,'start_spot_y':391, 'end_spot_x':3796,'end_spot_y':4232,'image_sort':17},
		'other ups label在左侧，多页的非第一页':{'degree':0,'format':[102,152],'save_width':800,'start_spot_x':1413,'start_spot_y':391, 'end_spot_x':3990,'end_spot_y':4157,'image_sort':18},
		'other ups label 逆时针90度，四周有空白':{'degree':3,'format':[102,152],'pdf_img_width':96,'save_width':800,'start_spot_x':722,'start_spot_y':366, 'end_spot_x':4632,'end_spot_y':6384,'image_sort':19},
	}
	
#png保存pdf	
def img_to_pdf(png_path,pdf_path,config):
	pdf = FPDF("P","mm",config['format'])
	pdf.add_page()
	pdf.set_font('Arial', 'B', 16)
	pdf.image(png_path,config.get('pdf_x',0),config.get('pdf_y',0),config.get('pdf_img_width',config['format'][0]))
	pdf.output(pdf_path)

#pdf转图片
def pdf_to_img(pdf_path,image_path):
	if os.path.exists(image_path):
		os.remove(image_path)
	pdfDoc = fitz.open(pdf_path)
	for pg in range(pdfDoc.page_count):
		page = pdfDoc[pg]
		rotate = int(0)
		# 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
		# 此处若是不做设置，默认图片大小为：792X612, dpi=96
		zoom_x = 9  # (1.33333333-->1056x816)   (2-->1584x1224)
		zoom_y = 9
		mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
		pix = page.get_pixmap(matrix=mat, alpha=False)
		pix._writeIMG(image_path,pg)  # 将图片写入指定的文件夹内
		

		#像素和毫米处理	 1mm = 3.78像素 210*297	
def png_deal(origin_png,config):
	# 旋转图片
	degree = config['degree']
	rotate_png(origin_png,degree)
	# 裁剪图片,缩放比例
	cut_png(origin_png,config)

#裁剪图片
def cut_png(origin_png,config):
	img = Image.open(origin_png) ## 打开chess.png文件，并赋值给img
	img.save(origin_png.replace('.png','-1.png'))
	if	config['end_spot_y']:
		region = img.crop((config['start_spot_x'],config['start_spot_y'],config['end_spot_x'],config['end_spot_y']))## 0,0表示要裁剪的位置的左上角坐标，50,50表示右下角。
	else:
		region = img
	(x,y) = region.size #读取图片尺寸（像素）
	x_s = config['save_width'] #定义缩小后的标准宽度 
	y_s = int(y * x_s / x) #基于标准宽度计算缩小后的高度
	out = region.resize((x_s,y_s),Image.Resampling.LANCZOS) #改变尺寸，保持图片高品质
	out.save(origin_png) ## 将裁剪下来的图片保存到 举例.png

def resize_png(origin_png):
	region = Image.open(origin_png) ## 打开chess.png文件，并赋值给img
	(x,y) = region.size #读取图片尺寸（像素）
	x_s = 250 #定义缩小后的标准宽度 
	y_s = int(y * x_s / x) #基于标准宽度计算缩小后的高度
	out = region.resize((x_s,y_s),Image.Resampling.LANCZOS) #改变尺寸，保持图片高品质
	out.save(origin_png) ## 将裁剪下来的图片保存到 举例.png
	return origin_png
	
#旋转图片	
def rotate_png(origin_png,degree):
	#img=cv2.imread(origin_png,1)
	img=cv2.imdecode(np.fromfile(origin_png, dtype=np.uint8), -1)
	for i in range(degree):
		img = np.rot90(img)
	dir = os.path.dirname(origin_png)
	origin_png_name = os.path.basename(origin_png)
	#cv2.imwrite(dir + '/_' + origin_png_name,img)
	cv2.imencode(".png",img)[1].tofile(dir + '/_' + origin_png_name)
	shutil.move(dir + '/_' + origin_png_name,origin_png)

#整体裁剪
def pdf_deal(file_name,tag_dir,need_create_dir,config):
	file_to_png = file_name[0:-4] + '.png'
	pdf_to_img(file_name,file_to_png)
	png_deal(file_to_png,config)
	# 生成pdf
	pdf_path = tag_dir + need_create_dir + '/' + os.path.basename(file_name)
	if not os.path.exists(os.path.dirname(pdf_path)):
			os.makedirs(os.path.dirname(pdf_path))
	img_to_pdf(file_to_png,pdf_path,config)
	
	return file_to_png


while True:
	event, values = window.read()
	progress_bar = window['progress_bar']
	window['progress_bar_text'].update(visible=False)
	window['progress_bar'].update(visible=False)
	window['progress_bar_num'].update(visible=False)
	window['img_deal_text'].update(visible=False)
	if event == sg.WIN_CLOSED:
		break
	elif event == "splitPdf":
		print('----start split ----')
		last_dir = os.path.basename(values['folder_path']);
		base_dir = values['folder_path'].replace(last_dir,'') + '/'
		dir = base_dir + last_dir
		tag_dir = base_dir + 'pySplit/' + str(get_now_datetime()) + '/'
		if not os.path.exists(os.path.dirname(tag_dir)):
			os.makedirs(os.path.dirname(tag_dir))
		file_list = recursion_dir_all_file(dir)
		for file_name in file_list:
			need_create_dir = os.path.dirname(file_name).replace(dir,'').replace('/','')
			if not os.path.exists(os.path.dirname(file_name)):
				os.makedirs(os.path.dirname(file_name))
			split_pdf(file_name,tag_dir,need_create_dir)
		print('----end split ----')
		sg.popup('啦啦啦','pdf自动拆分已完成！')
	elif event == 'cut_type':
		config = cut_type_config()[values['cut_type']]
		image_sort = config['image_sort']
		window['image'].update('./img/' + str(image_sort) + '-1.png')
		window['image1'].update('./img/' + str(image_sort) + '-2.png')
		window['image2'].update('./img/' + str(image_sort) + '-3.png')
		window['img_deal_0'].update('')
		window['img_deal_1'].update('')
		window['img_deal_2'].update('')
		progress_bar.UpdateBar(0)
	elif event == 'pdfDeal':
		#处理label
		window['progress_bar_text'].update(visible=True)
		window['progress_bar'].update(visible=True)
		window['progress_bar_num'].update(visible=True)
		window['img_deal_text'].update(visible=True)
		last_dir = os.path.basename(values['labels']);
		base_dir = values['labels'].replace(last_dir,'') + '/'
		tag_dir = base_dir + '/' + str(get_now_datetime()) + '/'
		config = cut_type_config()[values['cut_type']]
		pdf_dir = base_dir + last_dir
		file_list = recursion_dir_all_file(pdf_dir)
		i = 0
		bar_num = 0
		num = len(file_list)
		per_add = 1000 // num
		for file_name in file_list:
			bar_num = bar_num + per_add
			progress_bar.UpdateBar(bar_num)
			window['progress_bar_num'].update(str(bar_num/10) +'%')
			need_create_dir = os.path.dirname(file_name).replace(pdf_dir,'').replace('/','')
			if not os.path.exists(os.path.dirname(file_name)):
				os.makedirs(os.path.dirname(file_name))
			if os.path.splitext(file_name)[-1] != '.pdf':
				continue
			file_to_png = pdf_deal(file_name,tag_dir,need_create_dir,config)
			if i < 3:
				file_to_png = resize_png(file_to_png)
				tmp_name = 'img_deal_' + str(i)
				window[tmp_name].update(file_to_png)
				
			i = i+1
		window['progress_bar_num'].update('100%')
		progress_bar.UpdateBar(1000)
		sg.popup('自动裁剪label','---pdf已处理完成！---')
		
window.close()