#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import datetime
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMouseListener
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import ScrollBarOrientation  # 定数
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.util import XCloseListener
from com.sun.star.document import XDocumentEventListener
from com.sun.star.awt import Rectangle  # Struct
from com.sun.star.awt import PopupMenuDirection  # 定数
from com.sun.star.awt import XMenuListener
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	controller = doc.getCurrentController()  # コントローラの取得。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(controller, ctx, smgr, doc)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)  # EnhancedMouseClickHandler	
	doc.addDocumentEventListener(DocumentEventListener(enhancedmouseclickhandler))  # DocumentEventListener	
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, enhancedmouseclickhandler):
		self.args = enhancedmouseclickhandler
	def documentEventOccured(self, documentevent):
		enhancedmouseclickhandler = self.args
		if documentevent.EventName=="OnUnload":  
			source = documentevent.Source
			source.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
			source.removeDocumentEventListener(self)
	def disposing(self, eventobject):
		eventobject.Source.removeDocumentEventListener(self)
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, subj, ctx, smgr, doc):
		self.subj = subj
		self.args = ctx, smgr, doc
	def mousePressed(self, enhancedmouseevent):
		ctx, smgr, doc = self.args
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
					cellbackcolor = target.getPropertyValue("CellBackColor")  # セルの背景色を取得。
					if cellbackcolor==0x6666FF:  # 背景が青色の時。
						createDialog(ctx, smgr, doc, True)  # ノンモダルダイアログにする。	
						return False  # セル編集モードにしない。
					elif cellbackcolor==0xFFFF99:  # 背景が黄色の時。	
						createDialog(ctx, smgr, doc, False)  # モダルダイアログにする。		
						return False  # セル編集モードにしない。
		return True  # セル編集モードにする。
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # ドキュメントを閉じる時でも呼ばれない。
		self.subj.removeEnhancedMouseClickHandler(self)	
def createDialog(ctx, smgr, doc, flg):	
	frame = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = frame.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
	m = 6  # コントロール間の間隔
	grid = {"PositionX": m, "PositionY": m, "Width": 145, "Height": 100, "ShowColumnHeader": True, "ShowRowHeader": True}  # グリッドコントロールの基本プロパティ。
	label = {"PositionX": m, "Width": 45, "Height": 12, "Label": "Date and time: ", "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # ラベルフィールドコントロールの基本プロパティ。
	x = label["PositionX"]+label["Width"]  # ラベルフィールドコントロールの右端。
	textbox = {"PositionX": x, "Width": grid["PositionX"]+grid["Width"]-x, "Height": label["Height"], "VerticalAlign": MIDDLE}  # テクストボックスコントロールの基本プロパティ。
	button = {"PositionX": m, "Width": 30, "Height": label["Height"]+2, "PushButtonType": 0}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。
	controldialog =  {"PositionX": 100, "PositionY": 40, "Width": grid["PositionX"]+grid["Width"]+m, "Title": "Grid Example", "Name": "controldialog", "Step": 0, "Moveable": True}  # コントロールダイアログの基本プロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
	dialog, addControl = dialogCreator(ctx, smgr, controldialog)  # コントロールダイアログの作成。
	mouselister = MouseListener(doc, menuCreator(ctx, smgr))
	actionlistener = ActionListener()
	grid1 = addControl("Grid", grid, {"addMouseListener": mouselister})  # グリッドコントロールの取得。
	gridmodel = grid1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn()  # 列の作成。
	column0.Title = "Date"  # 列ヘッダー。
	column0.ColumnWidth = 60  # 列幅。
	gridcolumn.addColumn(column0)  # 列を追加。
	column1 = gridcolumn.createColumn()  # 列の作成。
	column1.Title = "Time"  # 列ヘッダー。
	column1.ColumnWidth = grid["Width"] - column0.ColumnWidth  #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 列を追加。	
	griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	now = datetime.now()  # 現在の日時を取得。
	griddata.addRow(0, (now.date().isoformat(), now.time().isoformat()))  # グリッドに行を挿入。
	y = grid["PositionY"] + grid["Height"] + m  # 下隣のコントロールのY座標を取得。
	label["PositionY"] = textbox["PositionY"] = y
	textbox["Text"] = now.isoformat().replace("T", " ")
	addControl("FixedText", label)
	addControl("Edit", textbox)  
	y = label["PositionY"] + label["Height"] + m  # 下隣のコントロールのY座標を取得。 
	button1, button2 = button.copy(), button.copy()
	button1["PositionY"] = button2["PositionY"]  = y
	button1["Label"] = "~Now"
	button2["Label"] = "~Close"
	button2["PushButtonType"] = 2  # CANCEL		
	button2["PositionX"] = grid["Width"] - button2["Width"]
	button1["PositionX"] = button2["PositionX"] - m - button1["Width"]
	addControl("Button", button1, {"setActionCommand": "now" ,"addActionListener": actionlistener})  
	addControl("Button", button2)  
	dialog.getModel().setPropertyValue("Height", button1["PositionY"]+button1["Height"]+m)  # コントロールダイアログの高さを設定。
	dialog.createPeer(toolkit, containerwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	if flg:  # ノンモダルダイアログにするとき。オートメーションでは動かない。
		dialogframe = showModelessly(ctx, smgr, frame, dialog)  
		dialogframe.addCloseListener(CloseListener(dialog, mouselister, actionlistener))  # CloseListener
	else:  # モダルダイアログにする。フレームに追加するとエラーになる。
		dialog.execute()  
		dialog.dispose()	
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, dialog, mouselister, actionlistener):
		self.args = dialog, mouselister, actionlistener
	def queryClosing(self, eventobject, getsownership):
		dialog, mouselister, actionlistener = self.args
		dialog.getControl("Grid1").removeMouseListener(mouselister)
		dialog.getControl("Button1").removeActionListener(actionlistener)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, doc, createMenu):
		menulistener = MenuListener()  # ポップアップメニューにつけるメニューリスナーを取得。
		items = ("~Cut", 0, {"setCommand": "cut"}),\
				("Cop~y", 0, {"setCommand": "copy"}),\
				("~Paste", 0, {"setCommand": "paste"}),\
				(),\
				("~Insert", 0, {"setCommand": "insert"}),\
				("~Delete", 0, {"setCommand": "delete"})
		popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		self.args = doc, popupmenu
	def mousePressed(self, mouseevent):
		doc, popupmenu = self.args
		source = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。
			selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				griddata = source.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
				rowdata = griddata.getRowData(source.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
				selection.setString(" ".join(rowdata))  # 選択セルに書き込む。
		elif mouseevent.PopupTrigger:  # 右クリックの時。
			if source.hasSelectedRows():  # 1つ以上の行が選択されている時。
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				popupmenu.execute(source.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向			
			
			
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)
class MenuListener(unohelper.Base, XMenuListener):
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		if cmd=="cut":
			pass
		elif cmd=="copy":
			pass
		elif cmd=="paste":
			pass
		elif cmd=="insert":
			pass
		elif cmd=="delete":
			pass
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)
def menuCreator(ctx, smgr):  #  メニューバーまたはポップアップメニューを作成する関数を返す。
	def createMenu(menutype, items, attr=None):  # menutypeはMenuBarまたはPopupMenu、itemsは各メニュー項目の項目名、スタイル、適用するメソッドのタプル、attrは各項目に適用する以外のメソッド。 
		if attr is None:
			attr = {}
		menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx) 
		for i, item in enumerate(items, start=1):  # 各メニュー項目について。
			if item:
				if len(item) > 2:  # タプルの要素が3以上のときは3番目の要素は適用するメソッドの辞書と考える。
					item = list(item)
					attr[i] = item.pop()  # メニュー項目のIDをキーとしてメソッド辞書に付け替える。
				menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。
			else:  # 空のタプルの時は区切り線と考える。
				menu.insertSeparator(i)  # ItemPos 
		if attr:  # メソッドの適用。
			for key, val in attr.items():  # keyはメソッド名あるいはメニュー項目のID。
				if isinstance(val, dict):  # valが辞書の時はkeyは項目ID。valはcreateMenu()の引数のitemsであり、itemsの３番目の要素にキーをメソッド名とする辞書が入っている。
					for method, arg in val.items():  # 辞書valのキーはメソッド名、値はメソッドの引数。
						if method in ("checkItem", "enableItem", "setCommand", "setHelpCommand", "setHelpText", "setTipHelpText"):  # 第1引数にIDを必要するメソッド。
							getattr(menu, method)(key, arg)
						else:
							getattr(menu, method)(arg)
				else:
					getattr(menu, key)(val) 
		return menu
	return createMenu
class ActionListener(unohelper.Base, XActionListener):
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		source = actionevent.Source  # ボタンコントロールが返る。
		context = source.getContext()  # コントロールダイアログが返ってくる。
		if cmd == "now":
			now = datetime.now()  # 現在の日時を取得。
			context.getControl("Edit1").setText(now.isoformat().replace("T", " "))  # テキストボックスコントロールに入力。
			grid1 = context.getControl("Grid1")  # グリッドコントロールを取得。
			griddata = grid1.getModel().getPropertyValue("GridDataModel")  # グリッドコントロールモデルからDefaultGridDataModelを取得。
			griddata.addRow(griddata.RowCount, (now.date().isoformat(), now.time().isoformat()))  # 新たな行を追加。
			accessiblecontext = grid1.getAccessibleContext()  # グリッドコントロールのAccessibleContextを取得。
			for i in range(accessiblecontext.getAccessibleChildCount()):  # 子要素をのインデックスを走査する。
				child = accessiblecontext.getAccessibleChild(i)  # 子要素を取得。
				if child.getAccessibleContext().getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
					if child.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
						child.setValue(child.getMaximum())  # 最大値にスクロールさせる。
						return
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションでは動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。	
	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。ｽﾍﾟｰｽは不可。
	parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
	dialog.setVisible(True)  # ダイアログを見えるようにする。   
	return frame  # フレームにリスナーをつけるときのためにフレームを返す。
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
		dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		control = None
		items, currentitemid = None, None
		if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
			if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
				items = props.pop("Items")
				if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
					currentitemid = props.pop("CurrentItemID")
		if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
			if controltype=="Grid":
				control = smgr.createInstanceWithContext("com.sun.star.awt.grid.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			else:	
				control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
			dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
		else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if items is not None:  # コントロールに追加されたRoadmapモデルにしかRoadmapアイテムは追加できない。
			for i, j in enumerate(items):  # 各Roadmapアイテムについて
				item = controlmodel.createInstance()
				item.setPropertyValues(("Label", "Enabled"), j)
				controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される
			if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
				controlmodel.setPropertyValue("CurrentItemID", currentitemid)
		if control is None:  # コントロールがまだインスタンス化されていないとき
			control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		if controltype=="Grid":
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.grid.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		else:	
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
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
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
