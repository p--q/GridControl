#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.beans import NamedValue
from com.sun.star.awt.PosSize import Y, WIDTH, POS, SIZE, POSSIZE
from com.sun.star.awt.ImageScaleMode import ISOTROPIC, ANISOTROPIC, NONE as ImageScaleMode_NONE
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK as ExecutableDialogResults_OK
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE
from com.sun.star.awt import XWindowListener, XItemListener, XActionListener, Rectangle
from com.sun.star.style.VerticalAlignment import MIDDLE
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	def wrapper(*args, **kwargs):
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
# @enableRemoteDebugging
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", Rectangle(100, 100, 530, 290)), NamedValue("FrameName", "ImageControlSample")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	frame = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
	window = frame.getContainerWindow()  # 新しいコンテナウィンドウを新しいフレームから取得。
	frame.setTitle("Image Viewer")  # フレームのタイトルを設定。
	docframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。
	actionlistener = ActionListener(ctx, smgr, frame)  # ボタンにつけるリスナー。コントロールコンテナはコントロールが追加されてから取得する。
	margin_horizontal = 20  # 水平マージン
	margin_vertical = 13  # 垂直マージン
	window_width = 537  # ウィンドウ幅
	window_height = 287  # ウィンドウの高さ
	line_height = 23  # Editコントロールやボタンコントロールなどの高さ
	buttonfilepick_width = 56  # ButtonFilePickボタンの幅。
	buttonclose_width = 90  # ButtonCloseボタンの幅。
	radiobutton_width = 120  # RadioButtonの幅。
	datalabel_width = 200  # FixedText2、FixedText3の幅。
	minwidth = margin_horizontal*2+radiobutton_width+buttonfilepick_width+5  # Edit1の幅>=1を幅の下限値とする。
	minheight = margin_vertical*5+line_height*2+1  # ImageControlの高さ>=1を高さの下限値とする。
	minsizes = minwidth, minheight  # コントロールが潰れてしまうと次のリサイズの計算がおかしくなるので下限値を設定する。
	editcontrol_Y = window_height-(margin_vertical+line_height)*2  # Edit1のY。頻用値を取得しておく。
	imagecontrolmargins = margin_horizontal*2, margin_vertical*5+line_height*2  # ウィンドウサイズの幅高さ-イメージコントロールの幅高さを取得。
	itemlistener = ItemListener(toolkit, window, imagecontrolmargins)
	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
	uno_path = pathsubstservice.getSubstituteVariableValue("$(prog)")  # fileurlでprogramフォルダへのパスが返ってくる。
	fileurl = "{}/intro.png".format(uno_path)  # 画像ファイルへのfileurl
	imageurl = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl))  # fileurlをシステム固有のパスに変換して正規化する。
	controlcontainer, addControl = controlcontainerCreator(ctx, smgr, {"PositionX": 0, "PositionY": 0, "Width": window_width, "Height": window_height, "BackgroundColor": 0xF0F0F0, "PosSize": POSSIZE})  # ウィンドウに表示させるコントロールコンテナを取得。BackgroundColor: -1は不可。
	imagecontrol1 = addControl("ImageControl", {"PositionX": margin_horizontal, "PositionY": margin_vertical, "Width": window_width-margin_horizontal*2, "Height": window_height-margin_vertical*5-line_height*2, "Border": 0, "ScaleImage": True, "ScaleMode": ISOTROPIC, "ImageURL": fileurl, "PosSize": POSSIZE})
	edit1 = addControl("Edit", {"PositionX": margin_horizontal+radiobutton_width+2, "PositionY": editcontrol_Y, "Width": window_width-margin_horizontal*2-radiobutton_width-buttonfilepick_width-4, "Height": line_height, "Text": imageurl, "PosSize": POSSIZE})
	button1 = addControl("Button", {"PositionX": window_width-margin_horizontal-buttonfilepick_width, "PositionY": editcontrol_Y, "Width": buttonfilepick_width, "Height": line_height, "Label": "~Browse", "PosSize": POSSIZE}, {"setActionCommand": "filepick" ,"addActionListener": actionlistener})  # PushButtonTypeは動かない。
	radiobutton1 = addControl("RadioButton", {"PositionX": margin_horizontal, "PositionY": editcontrol_Y, "Width": radiobutton_width, "Height": line_height, "Label": "~No Scaling", "PosSize": POSSIZE}, {"addItemListener": itemlistener})
	radiobutton2 = addControl("RadioButton", {"PositionX": margin_horizontal, "PositionY": editcontrol_Y+line_height+int((margin_vertical*2-line_height)/3), "Width": radiobutton_width, "Height": line_height, "Label": "~ISOTROPIC", "State": 1, "PosSize": POSSIZE}, {"addItemListener": itemlistener})
	radiobutton3 = addControl("RadioButton", {"PositionX": margin_horizontal, "PositionY": editcontrol_Y+(line_height+int((margin_vertical*2-line_height)/3))*2, "Width": radiobutton_width, "Height": line_height, "Label": "~ANISOTROPIC", "PosSize": POSSIZE}, {"addItemListener": itemlistener})
	imagecontrolrectangle = imagecontrol1.getPosSize()
	graphic = imagecontrol1.getModel().getPropertyValue("Graphic")  # イメージコントロールの画像を取得。
	graphicrectangle = graphic.getPropertyValue("SizePixel")  # 画像のオリジナルの大きさを取得。
	fixedtext1 = addControl("FixedText", {"PositionX": margin_horizontal+radiobutton_width+2, "PositionY":  editcontrol_Y+line_height+int((margin_vertical*2-line_height)/3), "Width": datalabel_width, "Height": line_height, "Label": "{} x {} px Image Size".format(graphicrectangle.Width, graphicrectangle.Height), "VerticalAlign": MIDDLE, "PosSize": POSSIZE})  # デフォルトでは"VerticalAlign": TOP、になっている。
	fixedtext2 = addControl("FixedText", {"PositionX": margin_horizontal+radiobutton_width+2, "PositionY": editcontrol_Y+(line_height+int((margin_vertical*2-line_height)/3))*2, "Width": datalabel_width, "Height": line_height, "Label": "{} x {} px Display Size".format(imagecontrolrectangle.Width, imagecontrolrectangle.Height), "VerticalAlign": MIDDLE, "PosSize": POSSIZE})
	button2 = addControl("Button", {"PositionX": window_width-margin_horizontal-buttonclose_width, "PositionY": window_height-margin_vertical-line_height, "Width": buttonclose_width, "Height": line_height, "Label": "~Close dialog", "PosSize": POSSIZE}, {"setActionCommand": "close" ,"addActionListener": actionlistener})  # PushButtonTypeは動かない。
	actionlistener.setControl(imagecontrol1, edit1, fixedtext1)  # getControl()で追加するコントロールが追加されてからコントロールコンテナを取得する。
	radiobuttons = radiobutton1, radiobutton2, radiobutton3
	itemlistener.setControl(imagecontrol1, radiobuttons)
	controlcontainer.createPeer(toolkit, window)  # ウィンドウにコントロールを描画。
	controlcontainer.setVisible(True)  # コントロールの表示。
	window.setVisible(True)  # ウィンドウの表示
	controls = controlcontainer, imagecontrol1, edit1, button1, button2, radiobutton1, radiobutton2, radiobutton3, fixedtext1, fixedtext2
	window.addWindowListener(WindowListener(controls, minsizes))  # setVisible(True)でも呼び出されるので、その後でリスナーを追加する。
class ItemListener(unohelper.Base, XItemListener):
	def __init__(self, toolkit, window, imagecontrolmargins):
		workarea = toolkit.getWorkArea()  # ディスプレイサイズの取得。ただしprimary displayだけ。IsMaximized()の大きさは同じプロセスではうまく取得できない。
		maxwidth = workarea.Width - 30  # ランチャーの幅を30pxと見積もる。
		maxheight = workarea.Height - 60  # ツールバーの高さを60pxと見積もる。
		self.window = window
		margin_width, margin_height = imagecontrolmargins
		self.consts = margin_width, margin_height, maxwidth, maxheight
	def setControl(self, imagecontrol, radiobuttons):
		self.imagecontrolmodel = imagecontrol.getModel()  # UnoControlImageの取得。
		controllabels = {}  # コントロールコンテナのUnoControlRadioButtonはLabelでしか判別できないので、Labelの辞書を作成しておく。
		imagescalemodes = ImageScaleMode_NONE, ISOTROPIC, ANISOTROPIC  # RadioButton1から3までに割り当てる。
		for radiobutton, imagescalemode in zip(radiobuttons, imagescalemodes):
			label = radiobutton.getModel().getPropertyValue("Label")
			controllabels[label] = imagescalemode
		self.controllabels = controllabels
# 	@enableRemoteDebugging
	def itemStateChanged(self, itemevent):  # コントロールコンテナのUnoControlRadioButtonはSourceから判断するしかない。
		window = self.window
		controlmodel = itemevent.Source.getModel()  # 発火させたコントロールモデルを取得。
		controllabel = controlmodel.getPropertyValue("Label")  # Labelプロパティを取得。
		self.imagecontrolmodel.setPropertyValue("ScaleMode", self.controllabels[controllabel])  # イメージコントロールのScaleModeプロパティを設定。
		if self.controllabels[controllabel]==ImageScaleMode_NONE:  # 縮尺しないとき
			graphic = self.imagecontrolmodel.getPropertyValue("Graphic")  # イメージコントロールの画像を取得。
			graphicrectangle = graphic.getPropertyValue("SizePixel")  # 画像のオリジナルの大きさを取得。
			margin_width, margin_height, maxwidth, maxheight = self.consts
			window_width = graphicrectangle.Width + margin_width  # 画像を全表示に必要なウィンドウの幅を取得。
			window_height = graphicrectangle.Height + margin_height  # 画像を全表示に必要なウィンドウの高さを取得。
			window_width = window_width if window_width<maxwidth else maxwidth
			window_height = window_height if window_height<maxheight else maxheight
			window.setPosSize(0, 0, window_width, window_height, SIZE)  # ウィンドウの大きさをイメージコントロールのサイズに合わす。
	def disposing(self, eventobject):
		pass
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, ctx, smgr, frame):
		self.ctx = ctx
		self.smgr = smgr
		self.frame = frame
		self.workurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # Ubuntuではホームフォルダ、Windows10ではドキュメントフォルダのfileurlが返る。
		self.simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)
	def setControl(self, imagecontrol, edit, fixedtext):
		self.imagecontrolmodel = imagecontrol.getModel()
		self.editcontrol = edit
		self.fixedtext = fixedtext
	def _createFilePicker(self, ctx, smgr):
		filters = {'WordPerfect Graphics': '*.wpg', 'SVM - StarView Meta File': '*.svm', 'PSD - Adobe Photoshop': '*.psd', 'EMF - Enhanced Meta File': '*.emf', 'PCD - Photo CD Base16': '*.pcd', 'PCD - Photo CD Base': '*.pcd', 'SGF - StarWriter SGF': '*.sgf', 'PGM - Portable Graymap': '*.pgm', 'SVG - Scalable Vector Graphics': '*.svg;*.svgz', 'PPM - Portable Pixelmap': '*.ppm', 'XBM - X Bitmap': '*.xbm', 'PBM - Portable Bitmap': '*.pbm', 'RAS - Sun Raster Image': '*.ras', 'WMF - Windows Metafile': '*.wmf', 'PCD - Photo CD Base4': '*.pcd', 'TGA - Truevision Targa': '*.tga', 'GIF - Graphics Interchange': '*.gif', 'Corel Presentation Exchange': '*.cmx', 'Adobe/Macromedia Freehand': '*.fh;*.fh1;*.fh2;*.fh3;*.fh4;*.fh5;*.fh6;*.fh7;*.fh8;*.fh9;*.fh10;*.fh11', 'CGM - Computer Graphics Metafile': '*.cgm', 'XPM - X PixMap': '*.xpm', 'MET - OS/2 Metafile': '*.met', 'DXF - AutoCAD Interchange Format': '*.dxf', 'JPEG - Joint Photographic Experts Group': '*.jpg;*.jpeg;*.jfif;*.jif;*.jpe', 'TIFF - Tagged Image File Format': '*.tif;*.tiff', 'PNG - Portable Network Graphic': '*.png', 'PCT - Mac Pict': '*.pct;*.pict', 'EPS - Encapsulated PostScript': '*.eps', 'BMP - Windows Bitmap': '*.bmp', 'PCX - Zsoft Paintbrush': '*.pcx'}  # 画像フィルターの辞書。
		filterall = "All Image Files"  # デフォルトで表示するフィルター名。
		template = FILEOPEN_SIMPLE
		try:  # 使えるのならFILEOPEN_PREVIEWを使う。
			from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_PREVIEW  # LibreOffice 5.3以上のみ
			template = FILEOPEN_PREVIEW
		except ImportError:
			pass
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (template,), ctx)
		filepicker.appendFilter("All Files", "*.*")  # すべてのファイルを表示させるフィルターを最初に追加。
		filepicker.appendFilter(filterall, ";".join(filters.values()))  # すべての画像ファイルを表示させるフィルターを2番目に追加。
		[filepicker.appendFilter(key, filters[key]) for key in sorted(filters.keys())]  # フィルターは追加された順に表示されるのでfiltersをキーでソートしてから追加している。
		filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルター名を設定。
		filepicker.setTitle("Insert Image")
		return filepicker
# 	@enableRemoteDebugging
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		if cmd == "filepick":
			filepicker = self._createFilePicker(self.ctx, self.smgr)
			systempath = self.editcontrol.getText().strip()  # Editコントロールのテキストを取得。システム固有形式のパスが入っているはず。
			if os.path.exists(systempath):  # パスが実存するとき
				if os.path.isfile(systempath):  # ファイルへのパスであればその親フォルダのパスを取得する。
					systempath = os.path.dirname(systempath)
				fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換する。
			else:
				fileurl = self.workurl  # 実存するパスが取得できない時はホームフォルダのfileurlを取得。
			filepicker.setDisplayDirectory(fileurl)  # 表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
			if filepicker.execute() == ExecutableDialogResults_OK:  # ファイル選択ダイアログを表示し、そのOKボタンがクリックされた時。
				fileurl = filepicker.getFiles()[0]  # ダイアログで選択されたファイルのパスを取得。fileurlのタプルで返ってくるので先頭の要素を取得。
				if self.simplefileaccess.exists(fileurl):  # fileurlが実存するとき
					self.imagecontrolmodel.setPropertyValue("ImageURL", fileurl)  # Imageコントロールに設定。
					systempath = unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステム固有形式に変換。
					self.editcontrol.setText(systempath)  # Editコントロールに表示。
					graphicrectangle = self.imagecontrolmodel.getPropertyValue("Graphic").getPropertyValue("SizePixel")  # 画像のオリジナルの大きさを取得。
					self.fixedtext.setText("{} x {} px Image Size".format(graphicrectangle.Width, graphicrectangle.Height))
		elif cmd == "close":
			self.frame.close(True)
	def disposing(self, eventobject):
		pass
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controls, minsizes):
		rectangle = controls[0].getPosSize()  # コントロールコンテナの位置と大きさを取得。なぜかwindow.getPosSize()では取得できない。
		self.oldwidth = rectangle.Width  # 変更前の幅を取得しておく。
		self.oldheight = rectangle.Height  # 変更前の高さを取得しておく。
		self.controls = controls
		self.minsizes = minsizes
# 	@enableRemoteDebugging
	def windowResized(self, windowevent):  # 変化分で計算する。コントロールが表示されないほど小さくされると次から表示がおかしくなる。
		minwidth, minheight = self.minsizes  # サイズ下限を取得。
		newwidth = windowevent.Width if windowevent.Width>minwidth else minwidth  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
		newheight = windowevent.Height if windowevent.Height>minheight else minheight  # 変更後のコントロールコンテナの高さを取得。サイズ下限より小さい時は下限値とする。
		self.diff_width = newwidth - self.oldwidth  # 幅変化分
		self.diff_height = newheight -self.oldheight  # 高さ変化分
		controlcontainer, imagecontrol1, edit1, button1, button2, radiobutton1, radiobutton2, radiobutton3, fixedtext1, fixedtext2 = self.controls  # 再計算するコントロールを取得。
		controlcontainer.setPosSize(0, 0, newwidth, newheight, SIZE)  # Flagsで変更する値のみ指定。変更しない値は0(でもなんでもよいはず)。
		self._applyDiff(fixedtext1, Y)
		self._applyDiff(fixedtext2, Y)
		self._applyDiff(imagecontrol1, SIZE)
		self._applyDiff(edit1, Y+WIDTH)
		self._applyDiff(radiobutton1, Y)
		self._applyDiff(radiobutton2, Y)
		self._applyDiff(radiobutton3, Y)
		self._applyDiff(button1, POS)
		self._applyDiff(button2, POS)
		imagecontrolrectangle = imagecontrol1.getPosSize()
		fixedtext2.setText("{} x {} px Display Size".format(imagecontrolrectangle.Width, imagecontrolrectangle.Height))
		self.oldwidth = newwidth  # 次の変更前の幅として取得。
		self.oldheight = newheight  # 次の変更前の高さとして取得。
	def _applyDiff(self, control, possize):  # 第2引数でウィンドウサイズの変化分のみ適用するPosSizeを指定。
		rectangle = control.getPosSize()  # 変更前のコントロールの位置大きさを取得。
		control.setPosSize(rectangle.X+self.diff_width, rectangle.Y+self.diff_height, rectangle.Width+self.diff_width, rectangle.Height+self.diff_height, possize)
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
def controlcontainerCreator(ctx, smgr, containerprops):  # コントロールコンテナと、それにコントロールを追加する関数を返す。まずコントロールコンテナモデルのプロパティを取得。UnoControlDialogElementサービスのプロパティは使えない。propsのキーにPosSize、値にPOSSIZEが必要。
	container = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールコンテナの生成。
	container.setPosSize(containerprops.pop("PositionX"), containerprops.pop("PositionY"), containerprops.pop("Width"), containerprops.pop("Height"), containerprops.pop("PosSize"))
	containermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コンテナモデルの生成。
	containermodel.setPropertyValues(tuple(containerprops.keys()), tuple(containerprops.values()))  # コンテナモデルのプロパティを設定。存在しないプロパティに設定してもエラーはでない。
	container.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、キーにPosSize、値にPOSSIZEが必要。attr: コントロールの属性。
		name = props.pop("Name") if "Name" in props else _generateSequentialName(controltype)  # サービスマネージャーからインスタンス化したコントロールはNameプロパティがないので、コントロールコンテナのaddControl()で名前を使うのみ。
		control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
		control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
		container.addControl(name, control)  # コントロールをコントロールコンテナに追加。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			control = container.getControl(name)  # コントロールコンテナに追加された後のコントロールを取得。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if controltype=="Grid":
			controlmodel = smgr.createInstance("com.sun.star.awt.grid.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		else:	
			controlmodel = smgr.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。

		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # コントロールの連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = container.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return container, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
	import sys
	from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()
