#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import datetime
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import XMenuListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.awt import PopupMenuDirection  # 定数
from com.sun.star.awt import Rectangle  # Struct
from com.sun.star.awt import ScrollBarOrientation  # 定数
from com.sun.star.document import XDocumentEventListener
from com.sun.star.util import XCloseListener
from com.sun.star.view.SelectionType import MULTI  # enum 
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.awt import MessageBoxButtons  # 定数
from com.sun.star.awt import MessageBoxResults  # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	controller = doc.getCurrentController()  # コントローラの取得。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(controller, ctx, smgr, doc)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)  # EnhancedMouseClickHandler	
	doc.addDocumentEventListener(DocumentEventListener(enhancedmouseclickhandler))  # DocumentEventListener。ドキュメントのリスナーの削除のため。	
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, enhancedmouseclickhandler):
		self.args = enhancedmouseclickhandler
	def documentEventOccured(self, documentevent):  # ドキュメントのリスナーを削除する。
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
	def mousePressed(self, enhancedmouseevent):  # try文を使わないと1回のエラーで以後動かなくなる。
		ctx, smgr, doc = self.args
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
					try:
						cellbackcolor = target.getPropertyValue("CellBackColor")  # セルの背景色を取得。
						if cellbackcolor==0x6666FF:  # 背景が青色の時。
							createDialog(ctx, smgr, doc, True)  # ノンモダルダイアログにする。	
							return False  # セル編集モードにしない。
						elif cellbackcolor==0xFFFF99:  # 背景が黄色の時。	
							createDialog(ctx, smgr, doc, False)  # モダルダイアログにする。		
							return False  # セル編集モードにしない。
					except:
						import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
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
	grid = {"PositionX": m, "PositionY": m, "Width": 100, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI, "VScroll": True}  # グリッドコントロールの基本プロパティ。
	textbox = {"PositionX": m, "PositionY": YHeight(grid, m), "Height": 12}  # テクストボックスコントロールの基本プロパティ。
	button = {"PositionY": textbox["PositionY"]-1, "Width": 23, "Height":textbox["Height"]+2, "PushButtonType": 2}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	controldialog =  {"PositionX": 100, "PositionY": 40, "Width": grid["PositionX"]+grid["Width"]+m, "Title": "Grid Example", "Name": "controldialog", "Step": 0, "Moveable": True}  # コントロールダイアログの基本プロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
	dialog, addControl = dialogCreator(ctx, smgr, controldialog)  # コントロールダイアログの作成。
	menulistener = MenuListener()
	menulistener.dialog = dialog  # メニューリスナーにダイアログを渡す。
	mouselistener = MouseListener(doc, menulistener, menuCreator(ctx, smgr))
	gridcontrol1 = addControl("Grid", grid, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn()  # 列の作成。
	column0.ColumnWidth = 50  # 列幅。
	gridcolumn.addColumn(column0)  # 列を追加。
	column1 = gridcolumn.createColumn()  # 列の作成。
	column1.ColumnWidth = grid["Width"] - column0.ColumnWidth  #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 列を追加。	
	griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	now = datetime.now()  # 現在の日時を取得。
	d = now.date().isoformat()
	t = now.time().isoformat().split(".")[0]
	griddata.addRow(0, (t, d))  # グリッドに行を挿入。
	textbox1, textbox2 = [textbox.copy() for dummy in range(2)]
	textbox1["Width"] = 34
	textbox1["Text"] = t
	textbox2["PositionX"] = XWidth(textbox1) 
	textbox2["Width"] = 42
	textbox2["Text"] = d
	addControl("Edit", textbox1)  
	addControl("Edit", textbox2, {"addMouseListener": mouselistener})  
	button["Label"] = "~Close"
	button["PositionX"] = XWidth(textbox2) 
	addControl("Button", button, {"addMouseListener": mouselistener})  
	dialog.getModel().setPropertyValue("Height", YHeight(button, m))  # コントロールダイアログの高さを設定。
	dialog.createPeer(toolkit, containerwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	if flg:  # ノンモダルダイアログにするとき。オートメーションでは動かない。
		dialogframe = showModelessly(ctx, smgr, frame, dialog)  
		args = doc, menulistener, mouselistener
		dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。
	else:  # モダルダイアログにする。フレームに追加するとエラーになる。
		dialog.execute()  
# 		mouselistener.gridpopupmenu.removeMenuListener(menulistener)
# 		mouselistener.editpopupmenu.removeMenuListener(menulistener)
# 		mouselistener.buttonpopupmenu.removeMenuListener(menulistener)		
		dialog.dispose()		
def XWidth(props, m=0):  # 左隣のコントロールからPositionXを取得。mは間隔。
	return props["PositionX"] + props["Width"] + m  	
def YHeight(props, m=0):  # 上隣のコントロールからPositionYを取得。mは間隔。
	return props["PositionY"] + props["Height"] + m
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):
		doc, menulistener, mouselistener = self.args
		dialog = menulistener.dialog
		gridcontrol = dialog.getControl("Grid1")	
		gridmodel = gridcontrol.getModel()  # グリッドコントロールモデルの取得。	
		datarows = [gridmodel.getRowData(i) for i in range(gridmodel.RowCount)]
		
# 		doc.getSheets()["config"]
		
		
		mouselistener.gridpopupmenu.removeMenuListener(menulistener)
		mouselistener.editpopupmenu.removeMenuListener(menulistener)
		mouselistener.buttonpopupmenu.removeMenuListener(menulistener)
		gridcontrol.removeMouseListener(mouselistener)
		dialog.getControl("Edit2").removeMouseListener(mouselistener)
		dialog.getControl("Button1").removeMouseListener(mouselistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, doc, menulistener, createMenu): 
		items = ("~Cut", 0, {"setCommand": "cut"}),\
			("Cop~y", 0, {"setCommand": "copy"}),\
			("~Paste Above", 0, {"setCommand": "pasteabove"}),\
			("P~aste Below", 0, {"setCommand": "pastebelow"}),\
			(),\
			("~Delete Selected Rows", 0, {"setCommand": "delete"})  # グリッドコントロールにつける右クリックメニュー。
		gridpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		self.gridpopupmenu = gridpopupmenu  # CloseListenerでも使う。
		items = ("~Now", 0, {"setCommand": "now"}),  # テキストボックスコントロールにつける右クリックメニュー。
		editpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		self.editpopupmenu = editpopupmenu  # CloseListenerでも使う。		
		items = ("~Resore", 0, {"setCommand": "restore"}),\
			(),\
			("~Add", 0, {"setCommand": "add"}),\
			("~Sort", 0, {"setCommand": "sort"})  # ボタンコントロールにつける右クリックメニュー。
		buttonpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		self.buttonpopupmenu = buttonpopupmenu  # CloseListenerでも使う。
		self.args = doc, menulistener
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。
		doc, menulistener = self.args
		name = mouseevent.Source.getModel().getPropertyValue("Name")
		if name=="Grid1":  # グリッドコントロールの時。
			gridpopupmenu = self.gridpopupmenu
			gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
			if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
					rowdata = griddata.getRowData(gridcontrol.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
					cellcursor = selection.getSpreadsheet().createCursorByRange(selection)  # 選択範囲のセルカーサーを取得。
					cellcursor.collapseToSize(len(rowdata), 1)  # (列、行)で指定。セルカーサーの範囲をrowdataに合せる。
					menulistener.undo = cellcursor, cellcursor.getDataArray()  # undoのためにセルカーサーとその値を取得する。
					cellcursor.setDataArray((rowdata,))  # セルカーサーにrowdataを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
			elif mouseevent.PopupTrigger:  # 右クリックの時。
				rowindex = gridcontrol.getRowAtPoint(mouseevent.X, mouseevent.Y)  # クリックした位置の行インデックスを取得。該当行がない時は-1が返ってくる。
				if rowindex>-1:  # クリックした位置に行が存在する時。
					flg = True  # Pasteメニューを表示させるフラグ。
					if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
						gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
						gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。
					rows = gridcontrol.getSelectedRows()  # 選択行インデックスを取得。
					rowcount = len(rows)  # 選択行数を取得。
					if rowcount>1 or not menulistener.rowdata:  # 複数行を選択している時または保存データがない時。
						flg = False  # Pasteメニューを表示しない。
					gridpopupmenu.enableItem(3, flg)  
					gridpopupmenu.enableItem(4, flg)  			
					pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
					gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向	
		elif name=="Edit2":  # テキストボックスコントロールの時。			
			if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。テキストボックスコントロールでは右クリックはカスタマイズ出来ない。
				editcontrol = mouseevent.Source  # テキストボックスコントロールを取得。
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				self.editpopupmenu.execute(editcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向						
		elif name=="Button1":  # ボタンコントロールの時。
			if mouseevent.PopupTrigger:  # 右クリックの時。
				flg = False  # Undoメニューを表示させるフラグ。
				if menulistener.undo:  # Undoデータがある時。
					cellcursor = menulistener.undo[0]  # Undoするセルカーサーを取得。
					activesheetname = doc.getCurrentController().getActiveSheet().getName()
					if activesheetname==cellcursor.getSpreadsheet().getName():  # Undoデータと同じシートの時。
						flg = True
				buttonpopupmenu = self.buttonpopupmenu
				buttonpopupmenu.enableItem(1, flg)  # Undoメニューを表示する。
				buttoncontrol = mouseevent.Source  # ボタンコントロールを取得。
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				buttonpopupmenu.execute(buttoncontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向					
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self):
		self.rowdata = None
		self.undo = None  # undo用データ。
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		dialog = self.dialog
		peer = dialog.getPeer()  # ピアを取得。
		toolkit = peer.getToolkit()  # ピアからツールキットを取得。 
		gridcontrol = dialog.getControl("Grid1")  # グリッドコントロールを取得。
		selectedrows = gridcontrol.getSelectedRows()  # 選択行インデックスのタプルを取得。
		griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
		if cmd in ("cut", "copy", "pasteabove", "pastebelow", "delete"):  # グリッドコントロールのコンテクストメニュー。
			if cmd=="cut":  # 選択行のデータを取得してその行を削除する。
				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
				[griddata.removeRow(r) for r in selectedrows]  # 選択行を削除。
			elif cmd=="copy":  # 選択行のデータを取得する。  
				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
			elif cmd=="pasteabove":  # 行を選択行の上に挿入。 
				insertRows(gridcontrol, griddata, selectedrows, 0, self.rowdata)
			elif cmd=="pastebelow":  # 空行を選択行の下に挿入。  
				insertRows(gridcontrol, griddata, selectedrows, 1, self.rowdata)
			elif cmd=="delete":  # 選択行を削除する。  
				msg = "Delete selected row(s)?"
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Delete", msg)
				if msgbox.execute()==MessageBoxResults.YES:					
					[griddata.removeRow(r) for r in selectedrows]  # 選択行を削除。
		elif cmd in ("add", "restore", "sort"):  # ボタンコントロールのコンテクストメニュー。
			if cmd=="add":
				dialog = self.dialog  # ダイアログを取得。
				t = dialog.getControl("Edit1").getText()
				d = dialog.getControl("Edit2").getText()			
				griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # グリッドコントロールモデルからDefaultGridDataModelを取得。
				if not selectedrows:  # 選択行がない時。
					selectedrows = griddata.RowCount-1,  # 最終行インデックスを選択していることにする。
				insertRows(gridcontrol, griddata, selectedrows, 1, ((t, d),))  # 選択行の下に行を挿入する。
			elif cmd=="sort":
				msg = "Sort in ascending order?"
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Sort", msg)
				if msgbox.execute()==MessageBoxResults.YES:				
					griddata.sortByColumn(0, True)
			elif cmd=="restore":
				cellcursor, datarows = self.undo  # datarowsは1行しかないはず。
				stringaddress = cellcursor.getPropertyValue("AbsoluteName").split(".")[1].replace("$", "")  # 前回入力した範囲の文字列アドレスを取得。
				current = " ".join(cellcursor.getDataArray()[0])
				restored = " ".join(datarows[0])
				msg = """Restore the Value of {}?
Current: {}
  After: {}""".format(stringaddress, current, restored)
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Undo", msg)
				if msgbox.execute()==MessageBoxResults.YES:
					cellcursor.setDataArray(datarows)
		elif cmd in ("now",):
			now = datetime.now()  # 現在の日時を取得。
			dialog.getControl("Edit2").setText(now.date().isoformat())  # テキストボックスコントロールに入力。
			dialog.getControl("Edit1").setText(now.time().isoformat().split(".")[0])  # テキストボックスコントロールに入力。			
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)
def insertRows(gridcontrol, griddata, selectedrows, position, datarows):  # positionは0の時は選択行の上に挿入、1で下に挿入。
	c = len(datarows)  # 行数を取得。
	griddata.insertRows(selectedrows[0]+position, ("", )*c, datarows)  # 行を挿入。
	gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
	gridcontrol.selectRow(selectedrows[0]+position)  # 挿入した行の最初の行を選択する。	
def menuCreator(ctx, smgr):  #  メニューバーまたはポップアップメニューを作成する関数を返す。
	def createMenu(menutype, items, attr=None):  # menutypeはMenuBarまたはPopupMenu、itemsは各メニュー項目の項目名、スタイル、適用するメソッドのタプルのタプル、attrは各項目に適用する以外のメソッド。
		if attr is None:
			attr = {}
		menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx)
		for i, item in enumerate(items, start=1):  # 各メニュー項目について。
			if item:
				if len(item) > 2:  # タプルの要素が3以上のときは3番目の要素は適用するメソッドの辞書と考える。
					item = list(item)
					attr[i] = item.pop()  # メニュー項目のIDをキーとしてメソッド辞書に付け替える。
				menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。ItemIdは1から始まり区切り線(欠番)は含まない。ItemPosは0から始まり区切り線を含む。
			else:  # 空のタプルの時は区切り線と考える。
				menu.insertSeparator(i-1)  # ItemPos
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
